import logging
from typing import Tuple, Any, Dict, List
from tablib import Dataset

from im_export.resources import InsureeResource

logger = logging.getLogger(__name__)


class InsureeImportExportService:
    supported_content_types = {
        'xls': 'application/vnd.ms-excel',
        'csv': 'text/csv',
        'json': 'application/json',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }

    class Strategy:
        INSERT = 'INSERT'
        UPDATE = 'UPDATE'
        INSERT_UPDATE = "INSERT_UPDATE"

    supported_strategies = (Strategy.INSERT, Strategy.UPDATE, Strategy.INSERT_UPDATE)

    def __init__(self, user):
        self._user = user
        self._resource = InsureeResource(user)

    def export_insurees(self, export_format: str = 'csv') -> Tuple[str, Any]:
        if export_format not in self.supported_content_types:
            raise ValueError(f'Non-supported export format: {export_format}')

        # All supported formats match Tablib attrs, to update if that's not valid anymore
        return self.supported_content_types[export_format], \
            getattr(self._resource.export(), export_format)

    def import_insurees(self, import_file, dry_run: bool = False, strategy: str = Strategy.INSERT) \
            -> Tuple[bool, Dict[str, int], List[str]]:

        if not import_file:
            return self._get_general_error('Missing import file')
        if strategy not in self.supported_strategies:
            return self._get_general_error(f'Non-supported strategy: {strategy}')

        # Other strategies are not supported for now
        if strategy in (InsureeImportExportService.Strategy.UPDATE, InsureeImportExportService.Strategy.INSERT_UPDATE):
            strategy = InsureeImportExportService.Strategy.INSERT
            logger.warning(f'Strategy {strategy} not currently supported, defaulting to {InsureeImportExportService.Strategy.INSERT}')

        try:
            if import_file.content_type == 'application/vnd.ms-excel':
                data_set = Dataset(headers=InsureeResource.insuree_headers).load(import_file.open(), 'xls')
            elif import_file.content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                data_set = Dataset(headers=InsureeResource.insuree_headers).load(import_file.open(), 'xlsx')
            elif import_file.content_type == 'application/json':
                data_set = Dataset(headers=InsureeResource.insuree_headers).load(import_file.open())
            else:
                data_set = Dataset(headers=InsureeResource.insuree_headers).load(import_file.read().decode())
        except Exception as e:
            return self._get_general_error('Failed to parse input file', e)

        try:
            data_set = self._resource.validate_and_sort_dataset(data_set)
        except Exception as e:
            return self._get_general_error('file validation failed', e)

        # in a rolled back transaction: the heads and their families exist for the members' rows
        dry_run_result = self._resource.import_data(data_set, dry_run=True, use_transactions=True)
        totals = self._get_totals_from_result(dry_run_result)
        errors = self._get_errors_from_result(dry_run_result)
        success = not dry_run_result.has_errors() and not dry_run_result.has_validation_errors()

        if not dry_run:
            if success:
                self._resource.import_data(data_set, dry_run=False, use_transactions=True)  # Actually import
                logger.info(f'Imported {totals["sent"]} insurees')
            else:
                logger.info(f'Failed to import {totals["sent"]} insurees, details: {totals}, errors: {errors}')

        return success, totals, errors

    @staticmethod
    def _get_totals_from_result(result):
        return {
            'sent': result.total_rows,
            'created': result.totals['new'],
            'updated': result.totals['update'],
            'deleted': result.totals['delete'],
            'skipped': result.totals['skip'],
            'invalid': result.totals['invalid'],
            'failed': result.totals['error']
        }
        
    @staticmethod
    def _get_general_error(*args):
        errors = []
        for arg in args:
            errors.append(arg.message if hasattr(arg, 'message') else str(arg))
        totals = {'sent': 0, 'created': 0, 'updated': 0, 'deleted': 0, 'skipped': 0, 'invalid': 0, 'failed': 0}
        success = False

        return success, totals, errors

    @staticmethod
    def _get_errors_from_result(result):
        errors = []
        if result.has_validation_errors():
            for invalid_row in result.invalid_rows:
                errors.append(f"row ({invalid_row.number}) - {invalid_row.error.messages}")
        if result.has_errors():
            for index, row_error in result.row_errors():
                for error in row_error:
                    errors.append(f"row ({index}) - {error.error}")
        return errors
