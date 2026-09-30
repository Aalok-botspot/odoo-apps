# -*- coding: utf-8 -*-
# Part of SmartImport AI. See LICENSE file for full copyright and licensing details.
import base64
import io
import logging
import re

from dateutil import parser as date_parser

from odoo import _, api, fields, models
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import pandas as pd
    from thefuzz import fuzz, process
except ImportError:  # Odoo also checks 'external_dependencies' at install time
    pd = fuzz = process = None
    _logger.warning("SmartImport AI needs: pip install pandas thefuzz openpyxl")

# A column is auto-mapped only when its fuzzy score is strictly above this value.
MATCH_THRESHOLD = 70

# Technical/audit fields that should never be filled from a file.
SKIP_FIELDS = {
    'id', 'create_uid', 'create_date', 'write_uid', 'write_date',
    'display_name', '__last_update',
}
# Field types that cannot be sensibly filled from a single spreadsheet cell.
UNSUPPORTED_TYPES = {
    'one2many', 'binary', 'json', 'properties', 'properties_definition',
    'reference', 'many2one_reference',
}
TRUE_VALUES = {'1', 'true', 'yes', 'y', 't', 'x', 'on', 'checked'}
FALSE_VALUES = {'0', 'false', 'no', 'n', 'f', 'off', 'unchecked'}
MAX_LOGGED_ERRORS = 50


class SmartImportWizard(models.TransientModel):
    _name = 'smart.import.wizard'
    _description = 'SmartImport AI Wizard'

    model_id = fields.Many2one(
        'ir.model', string='Import Into', required=True, ondelete='cascade',
        domain=[('transient', '=', False)],
        help="The Odoo model (Contacts, Products, Sales Orders...) that will receive the records.")
    res_model = fields.Char(
        related='model_id.model', string='Technical Model Name', readonly=True)
    file_data = fields.Binary(string='Excel / CSV File', required=True, attachment=False)
    file_name = fields.Char(string='File Name')
    skip_errors = fields.Boolean(
        string='Skip Errors', default=True,
        help="If a row fails (bad data, missing required value...), skip it and continue "
             "with the remaining rows. If disabled, the whole import is cancelled on the first error.")
    auto_clean = fields.Boolean(
        string='Auto-Clean Data', default=True,
        help="Trim spaces, ignore empty cells/NaN and tolerate formats like '1,200.50' or 'Rs 500'.")
    dayfirst = fields.Boolean(
        string='Dates are Day-First (DD/MM/YYYY)', default=False,
        help="Enable when ambiguous dates such as 03/04/2025 mean 3 April instead of March 4.")

    # ------------------------------------------------------------------
    # Defaults
    # ------------------------------------------------------------------
    @api.model
    def default_get(self, fields_list):
        """Allow other buttons/actions to pre-select the model with context 'default_res_model'."""
        res = super().default_get(fields_list)
        model_name = self.env.context.get('default_res_model')
        if 'model_id' in fields_list and not res.get('model_id') and model_name:
            model_id = self.env['ir.model']._get_id(model_name)
            if model_id:
                res['model_id'] = model_id
        return res

    # ------------------------------------------------------------------
    # Main entry point
    # ------------------------------------------------------------------
    def action_analyze_and_import(self):
        self.ensure_one()
        if not pd or not process:
            raise UserError(_("Python libraries missing. Please run: pip install pandas thefuzz openpyxl"))
        if not self.file_data:
            raise UserError(_("Please upload an Excel (.xlsx) or CSV file first."))

        model = self._get_target_model()                       # a) target model
        df = self._read_dataframe()                            # b) file -> DataFrame
        odoo_fields = self._get_importable_fields(model)       # c) fields_get()
        mapping = self._build_mapping(list(df.columns), odoo_fields)   # d) fuzzy mapping
        if not mapping:
            raise UserError(_(
                "No column could be matched to a field of '%(model)s'. "
                "Please check that the first row of your file contains column headers.",
                model=self.model_id.name))
        _logger.info("SmartImport AI mapping for %s: %s", model._name, mapping)

        imported, skipped, errors = self._import_rows(model, df, mapping, odoo_fields)

        message = _("%(ok)s records imported successfully, %(bad)s records skipped.",
                    ok=imported, bad=skipped)
        message += " " + _("(%(mapped)s of %(total)s columns were auto-mapped.)",
                           mapped=len(mapping), total=len(df.columns))
        if errors:
            message += " " + _("First problem: %s", errors[0])
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _("SmartImport AI"),
                'message': message,
                'type': 'warning' if skipped else 'success',
                'sticky': bool(skipped),
                'next': {'type': 'ir.actions.act_window_close'},
            },
        }

    # ------------------------------------------------------------------
    # Step helpers
    # ------------------------------------------------------------------
    def _get_target_model(self):
        """Return the target model, using the current user's rights (no sudo)."""
        model_name = self.res_model
        if not model_name or model_name not in self.env:
            raise UserError(_("Please select a valid model to import into."))
        model = self.env[model_name]
        if model._abstract or model._transient:
            raise UserError(_("Records cannot be imported into abstract or transient models."))
        model.check_access('create')   # raises AccessError if the user may not create records
        return model

    @staticmethod
    def _decode_text(raw):
        for encoding in ('utf-8-sig', 'cp1252'):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw.decode('latin-1')

    def _read_dataframe(self):
        """Read the uploaded .csv / .xlsx file into a pandas DataFrame (all cells as text)."""
        raw = base64.b64decode(self.file_data)
        name = (self.file_name or '').lower()
        if name.endswith('.xls'):
            raise UserError(_("Old .xls files are not supported. Please save the file as .xlsx or .csv."))
        is_excel = name.endswith(('.xlsx', '.xlsm')) or (not name.endswith('.csv') and raw[:2] == b'PK')
        try:
            if is_excel:
                df = pd.read_excel(io.BytesIO(raw), engine='openpyxl', dtype=str)
            else:
                text = self._decode_text(raw)
                try:  # sniff the delimiter (comma, semicolon, tab...)
                    df = pd.read_csv(io.StringIO(text), sep=None, engine='python',
                                     dtype=str, keep_default_na=False)
                except Exception:
                    df = pd.read_csv(io.StringIO(text), sep=',', dtype=str, keep_default_na=False)
        except Exception as exc:
            raise UserError(_("The file could not be read: %s", self._error_text(exc)))

        # Clean headers, drop unnamed columns and fully empty rows.
        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, [not col.startswith('Unnamed:') for col in df.columns]]
        df = df.dropna(how='all')
        if df.empty:
            raise UserError(_("The file does not contain any data rows."))
        return df

    @staticmethod
    def _get_importable_fields(model):
        """Storable, writable fields of the model, using Odoo's fields_get()."""
        meta = model.fields_get(
            attributes=['string', 'type', 'store', 'readonly', 'required', 'selection', 'relation'])
        return {
            name: attrs for name, attrs in meta.items()
            if name not in SKIP_FIELDS
            and attrs.get('type') not in UNSUPPORTED_TYPES
            and attrs.get('store')
            and not attrs.get('readonly')
        }

    @staticmethod
    def _normalize(text):
        """Lowercase and turn separators into spaces: 'Customer_Name / Ref' -> 'customer name ref'."""
        return re.sub(r'[\W_]+', ' ', str(text).lower()).strip()

    def _build_mapping(self, columns, odoo_fields):
        """Fuzzy-match file headers to Odoo fields. Returns {column_header: field_name}."""
        choices = {}
        for fname in odoo_fields:                       # technical names win on conflicts
            choices.setdefault(self._normalize(fname), fname)
        for fname, attrs in odoo_fields.items():        # then human labels
            choices.setdefault(self._normalize(attrs['string']), fname)
        choice_keys = list(choices)

        candidates = []                                 # (score, column, field_name)
        for col in columns:
            norm = self._normalize(col)
            if not norm:
                continue
            if norm in choices:                         # exact match, no fuzzy needed
                candidates.append((100, col, choices[norm]))
                continue
            hit = process.extractOne(norm, choice_keys, scorer=fuzz.WRatio)
            if hit and hit[1] > MATCH_THRESHOLD:
                candidates.append((hit[1], col, choices[hit[0]]))

        # If two columns want the same field, the best score wins.
        candidates.sort(key=lambda item: -item[0])
        used_fields, best = set(), {}
        for score, col, fname in candidates:
            if fname in used_fields:
                _logger.info("SmartImport AI: column '%s' ignored, field '%s' already mapped", col, fname)
                continue
            used_fields.add(fname)
            best[col] = fname
        return {col: best[col] for col in columns if col in best}   # keep file order

    # ------------------------------------------------------------------
    # Row processing
    # ------------------------------------------------------------------
    def _import_rows(self, model, df, mapping, odoo_fields):
        """Create one record per row. Returns (imported, skipped, error_messages)."""
        # Faster imports: no chatter tracking / creation logs / auto-subscription.
        model = model.with_context(
            tracking_disable=True, mail_create_nolog=True, mail_create_nosubscribe=True)
        imported = skipped = 0
        errors = []
        cache = {}   # avoids repeating the same many2one lookups

        for index, row in df.iterrows():
            row_number = int(index) + 2   # +1 for the header row, +1 because Excel starts at 1
            try:
                vals = self._prepare_vals(row, mapping, odoo_fields, cache)
                if not vals:
                    continue              # completely blank row
                if self.skip_errors:
                    # The savepoint rolls back ONLY this row if it fails,
                    # so rows imported before are kept.
                    with self.env.cr.savepoint():
                        model.create(vals)
                else:
                    model.create(vals)
                imported += 1
            except Exception as exc:
                text = _("Row %(row)s: %(error)s", row=row_number, error=self._error_text(exc))
                if not self.skip_errors:
                    # Raising cancels the whole transaction: all-or-nothing.
                    raise UserError(_("Import cancelled. %s", text))
                skipped += 1
                _logger.warning("SmartImport AI skipped %s", text)
                if len(errors) < MAX_LOGGED_ERRORS:
                    errors.append(text)
        return imported, skipped, errors

    def _prepare_vals(self, row, mapping, odoo_fields, cache):
        vals = {}
        for col, fname in mapping.items():
            raw = row[col]
            if self._is_empty(raw):
                continue                  # empty cells / NaN never crash the import
            try:
                vals[fname] = self._convert(raw, odoo_fields[fname], cache)
            except Exception as exc:
                raise ValueError(_("column '%(col)s': %(error)s", col=col, error=self._error_text(exc)))
        return vals

    @staticmethod
    def _is_empty(value):
        if value is None:
            return True
        try:
            if pd.isna(value):
                return True
        except (TypeError, ValueError):
            pass
        return not str(value).strip()

    @staticmethod
    def _error_text(exc):
        text = ' '.join(str(exc).split()) or exc.__class__.__name__
        return text[:300]

    # ------------------------------------------------------------------
    # Value conversion (text cell -> Odoo field value)
    # ------------------------------------------------------------------
    def _convert(self, raw, meta, cache):
        value = str(raw)
        if self.auto_clean:
            value = value.strip()
        ftype = meta['type']

        if ftype in ('char', 'text', 'html'):
            return value
        if ftype in ('float', 'monetary'):
            return self._to_float(value)
        if ftype == 'integer':
            number = self._to_float(value)
            if number != int(number):
                raise ValueError(_("'%s' is not a whole number", value))
            return int(number)
        if ftype == 'boolean':
            lowered = value.lower()
            if lowered in TRUE_VALUES:
                return True
            if lowered in FALSE_VALUES:
                return False
            raise ValueError(_("'%s' is not a valid yes/no value", value))
        if ftype == 'date':
            return date_parser.parse(value, dayfirst=self.dayfirst).date()
        if ftype == 'datetime':
            return date_parser.parse(value, dayfirst=self.dayfirst).replace(tzinfo=None)
        if ftype == 'selection':
            lowered = value.lower()
            for key, label in meta.get('selection') or []:
                if lowered in (str(key).lower(), str(label).lower()):
                    return key
            raise ValueError(_("'%s' is not an allowed choice", value))
        if ftype == 'many2one':
            return self._lookup_record(meta['relation'], value, cache)
        if ftype == 'many2many':
            names = [part.strip() for part in re.split(r'[;,|]', value) if part.strip()]
            ids = [self._lookup_record(meta['relation'], name, cache) for name in names]
            return [(6, 0, ids)]
        raise ValueError(_("field type '%s' is not supported", ftype))

    def _to_float(self, value):
        try:
            return float(value)
        except ValueError:
            if not self.auto_clean:
                raise ValueError(_("'%s' is not a number", value))
            # Remove currency symbols / thousand separators: 'Rs 1,200.50' -> 1200.50
            cleaned = re.sub(r'[^0-9.,\-]', '', value).replace(',', '')
            try:
                return float(cleaned)
            except ValueError:
                raise ValueError(_("'%s' is not a number", value))

    def _lookup_record(self, relation, name, cache):
        """Find the record of `relation` whose name matches `name` (case-insensitive)."""
        key = (relation, name.lower())
        if key in cache:
            return cache[key]
        comodel = self.env[relation]
        rec_name = comodel._rec_name or 'name'
        if rec_name not in comodel._fields:
            raise ValueError(_("cannot look up records of '%s' by name", relation))
        records = comodel.search([(rec_name, 'ilike', name)], limit=20)
        exact = records.filtered(lambda rec: (rec[rec_name] or '').strip().lower() == name.lower())
        if exact:
            record = exact[0]
        elif len(records) == 1:
            record = records
        elif records:
            raise ValueError(_("'%(name)s' matches several records of %(model)s", name=name, model=relation))
        else:
            raise ValueError(_("'%(name)s' was not found in %(model)s", name=name, model=relation))
        cache[key] = record.id
        return record.id
