import logging
from typing import Tuple, Any, Dict, List
from tablib import Dataset

from im_export.resources import InsureeResource
import openpyxl
from decimal import Decimal
from invoice.models import Invoice, PaymentInvoice, DetailPaymentInvoice, InvoiceEvent
from insuree.models import Insuree, Family, ResidenceEnvironment, HousingType, MutualInsuranceCoverage, NoDisability, NonDisablingDisease
from insuree.services import FamilyService, InsureeService
from contribution.models import Premium
from policy.models import Policy
from policy.services import PolicyService
from core.models import Officer
from datetime import datetime, timedelta
from uuid import uuid4
from django.contrib.contenttypes.models import ContentType
from contribution.services import update_or_create_premium
from collections import defaultdict
from location.models import Location
from django.db.models import Q
from datetime import datetime as py_datetime, date as py_date
from core.datetimes.shared import datetimedelta
from contribution_plan.models import ContributionPlan
from django.db import transaction 
from django.utils import timezone
from datetime import timedelta
from core.utils import TimeUtils
from product.models import Product
import os
from pathlib import Path
import copy
from policy.services import update_insuree_policies
from policy.services import policy_status_premium_paid
from dateutil.relativedelta import relativedelta
from invoice.services import InvoiceService
from invoice.services.invoiceLineItem import InvoiceLineItemService
import calendar
from policy.apps import CALCULATION_RULES
from policyholder.models import PolicyHolder
from uuid import UUID

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

        dry_run_result = self._resource.import_data(data_set, dry_run=True)  # Test the data import
        totals = self._get_totals_from_result(dry_run_result)
        errors = self._get_errors_from_result(dry_run_result)
        success = not dry_run_result.has_errors() and not dry_run_result.has_validation_errors()

        if not dry_run:
            if success:
                self._resource.import_data(data_set, dry_run=False)  # Actually import
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


class FamilyImportExportService:
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
        # self._resource = InsureeResource(user)

    def export_families(self, export_format: str = 'csv') -> Tuple[str, Any]:
        if export_format not in self.supported_content_types:
            raise ValueError(f'Non-supported export format: {export_format}')

        # All supported formats match Tablib attrs, to update if that's not valid anymore
        return self.supported_content_types[export_format], \
            getattr(self._resource.export(), export_format)

    def import_families(self, import_file, dry_run: bool = False, strategy: str = Strategy.INSERT) \
            -> Tuple[bool, Dict[str, int], List[str]]:

        if not import_file:
            return InsureeImportExportService._get_general_error('Missing import file')
        if strategy not in self.supported_strategies:
            return InsureeImportExportService._get_general_error(f'Non-supported strategy: {strategy}')

        # Other strategies are not supported for now
        if strategy in (InsureeImportExportService.Strategy.UPDATE, InsureeImportExportService.Strategy.INSERT_UPDATE):
            strategy = InsureeImportExportService.Strategy.INSERT
            logger.warning(f'Strategy {strategy} not currently supported, defaulting to {InsureeImportExportService.Strategy.INSERT}')

        family_headers = ['Identification', 'Etatmatrimonial', 'Membresménage',
       'Nom&prénom_membresménag', 'Sexe', 'Lien de parenté', "Pièce d'identité",
       'NIN', 'Jour de naissance', 'Mois de naissance', 'Année de naissance', 'Âge',
       'Formation ou non', 'Types de formation', 'Maladie invalidante Non',
       'Handicap Non', 'Couverture_Assurance_Mutuelle',
       'Catégories_professionnelles', 'Tailleménages', 'Scores_taille_des_ménages',
       'Revenus', 'Scores_revenus', 'Scores_types_habitation',
       'Scores_totaux_catégorisation', 'Cotisationsfamilles', 'Cotisations_famille_chef_famille',
       'île', 'milieu de résidence', 'District_sanitaire', 'Commune', 'Localité',
       'Taillefamille', 'Autresménage', 'Féminin', 'Masculin',
       'Cotisationsautresménages', 'Cotisations_totales_ménages',
       'PartsGouv_&_PTFCotisations Familles',
       'PartsGouv_&_PTFCotisations_Autresménages_démunis',
       'Parts Gouv & PTFCotisations_Autresménages_Vulnérables',
       'Partstotaux_Gvt&PTF']
        # Education, Relations, Profession, changer les chiffres de income levels
        try:
            if import_file.content_type == 'application/vnd.ms-excel':
                data_set = Dataset(headers=family_headers).load(import_file.open(), 'xls')
            elif import_file.content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                data_set = Dataset(headers=family_headers).load(import_file.open(), 'xlsx')
            elif import_file.content_type == 'application/json':
                data_set = Dataset(headers=family_headers).load(import_file.open())
            else:
                data_set = Dataset(headers=family_headers).load(import_file.read().decode())
        except Exception as e:
            return InsureeImportExportService._get_general_error('Failed to parse input file', e)

        non_disabling_dict = {
            1:1,
            0:2
        }

        no_disability_dict = {
            1:1,
            0:2
        }

        mutual_insurance_coverange_dict={
            1:2,
            0:1
        }

        gender_dict = {
            1: "M",
            2: "F",
            3: "O"
        }
        
        professional_situations = {
            0: "Sans profession",
            1: "Agriculteur exploitant/Pêcheur/Éleveur/ou Artisan",
            2: "Commerçants et assimilés",
            3: "Chef d'entreprise",
            4: "Professionnel libéral",
            5: "Enseignant",
            6: "Enseignant chercheur",
            7: "Médecin",
            8: "Autre professionnel de Santé",
            9: "Autorité de l'admin. publique",
            10: "Employé de l'administration publique",
            11: "Employé administratif ou commercial d'entreprise",
            12: "Ingénieur et Technicien d'entreprise",
            13: "Ingénieur et Technicien travaillant à compte propre",
            14: "Clergé, religieux",
            15: "Policier et militaire",
            16: "Personnel de services directs aux particuliers",
            17: "Artiste",
            18: "Ouvrier",
            19: "Retraité"
        }
        directory = "./FamilyImports"
        file_path = False

        target_dir = Path(directory)
        os.makedirs(target_dir, exist_ok=True)
        
        # Get original filename and extension
        original_name, ext = os.path.splitext(import_file.name)
        
        # Add current date to filename
        current_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        new_filename = f"{original_name}_{current_date}{ext}"
        
        # Create full file path
        file_path = os.path.join(target_dir, new_filename)
        
        # Save the file
        with open(file_path, 'wb+') as destination:
            for chunk in import_file.chunks():
                destination.write(chunk)
        grouped = defaultdict(list)
        today = py_datetime.today()
        for row in data_set.dict:
            grouped[row['Identification']].append(row)
        try:
            familycreated = 0
            with transaction.atomic():
                for identification, rows in grouped.items():
                    logger.info(f"Memmbers for Identification = {identification}:")
                    if identification:
                        policies = []
                        # Mise par Ordre de membre_de_menage ascendant en commencant par le chef de famille
                        sorted_members = sorted(rows, key=lambda x: int(x['Membresménage']) if x['Membresménage'] not in [None, ""] else -1)
                        parent_family = None
                        sub_family = None
                        number_count = 0
                        for r in sorted_members:
                            number_count += 1
                            logger.info(parent_family)
                            yob = r.get("Année de naissance")
                            if yob is None or yob == "-999999999" or len(str(yob)) != 4 or yob == "":
                                if r.get("Lien de parenté") != "" and r.get("Lien de parenté") is not None:
                                    if int(r.get("Lien de parenté")) == 3:
                                        years = 10
                                    elif int(r.get("Lien de parenté")) in [1, 2, 25]:
                                        years = 30
                                    elif int(r.get("Lien de parenté")) in [21, 22, 23, 24]:
                                        years = 40
                                    else:
                                        years = 50
                                    yob = str(today - datetimedelta(
                                        years=years
                                    )).split(" ")[0].split("-")[0]
                                else:
                                    yob = str(today - datetimedelta(
                                        years=50
                                    )).split(" ")[0].split("-")[0]
                            if r.get("Mois de naissance") is None or str(r.get("Mois de naissance")) in ["-999999999", ""]:
                                mob = "01"
                            else:
                                mob = str(r.get("Mois de naissance"))
                                mob = mob.zfill(2)# 1 becomes 01
                                if len(mob) != 2:
                                    mob = "01"
                            if r.get("Jour de naissance") is None or str(r.get("Jour de naissance")) in ["-999999999", ""]:
                                dob = "01"
                            else:
                                dob = str(r.get("Jour de naissance"))
                                dob = dob.zfill(2) # 1 becomes 01
                                if len(dob) != 2:
                                    dob = "01"
                            village = r.get("Localité")
                            if village == "" or village is None:
                                village = 1
                            current_village_id = Location.objects.filter(
                                Q(code=village) | Q(name=village)).filter(
                                    validity_to__isnull=True, type='V').first()
                            if not current_village_id:
                                current_village_id = Location.objects.filter(
                                    validity_to__isnull=True, type='V').first()
                            current_village_id = current_village_id.id
                            current_gender = gender_dict.get(3)
                            if r.get("Sexe") is not None and int(r.get("Sexe")) in [1, 2]:
                                current_gender = gender_dict.get(int(r.get("Sexe")))
                            nin_ok = False
                            nin = False
                            if r.get("NIN") is not None and r.get("NIN") != "":
                                nin = str(r.get("NIN")).replace(" ", "").replace("#", "")
                                if len(nin) == 7 or (len(nin)==9 and nin.startswith("UG")):
                                    nin_ok = True
                            card_issued = True
                            if r.get("Pièce d'identité") is not None and r.get("Pièce d'identité") != "":
                                if not nin_ok or int(r.get("Pièce d'identité")) != 1:
                                    card_issued = False
                            marital = r.get("Etatmatrimonial")
                            if marital is not None and marital != "":
                                my_dict = {
                                    "1": "M",
                                    "3": "D",
                                    "4": "W",
                                    "5": "S",
                                    "2": "P"
                                }
                                marital = my_dict.get(str(marital))
                            else:
                                marital = False
                            head_insuree_data = {
                                "last_name": r.get("Nom&prénom_membresménag") if r.get("Nom&prénom_membresménag")\
                                    not in [None, ""] else " ",
                                "other_names": " ",
                                "gender_id": current_gender,
                                "dob": str(yob) + "-" + str(mob) + "-" + str(dob),
                                "current_village_id": current_village_id,
                                "card_issued": card_issued,
                                "audit_user_id": self._user._u.id
                            }

                            non_disabling_code = r.get("Maladie invalidante Non")
                            if non_disabling_code is not None and str(non_disabling_code).isdigit():
                                head_insuree_data["non_disabling_disease_id"] = non_disabling_dict.get(int(non_disabling_code))

                            no_disability_code = r.get("Handicap Non")
                            if no_disability_code is not None and str(no_disability_code).isdigit():
                                head_insuree_data["no_disability_id"] = no_disability_dict.get(int(no_disability_code))

                            coverage_code = r.get("Couverture_Assurance_Mutuelle")
                            if coverage_code is not None and str(coverage_code).isdigit():
                                head_insuree_data["mutual_insurance_coverage_id"] = mutual_insurance_coverange_dict.get(int(coverage_code))

                            environment_code = r.get("milieu de résidence")
                            if environment_code is not None and str(environment_code).isdigit():
                                head_insuree_data["residence_environment_id"] = int(environment_code)

                            housing_type_code = r.get("Scores_types_habitation")
                            if housing_type_code is not None and str(housing_type_code).isdigit():
                                head_insuree_data["housing_type_id"] = housing_type_code

                            if marital:
                                head_insuree_data["marital"] = marital
                            if nin is not False and nin_ok:
                                head_insuree_data["passport"] = nin
                            if r.get("Revenus") is not None and r.get("Revenus") != "":
                                head_insuree_data["income_level_id"] = int(r.get("Revenus"))+1 #+1 parce que les revenus
                                # en BD commencent a 1 et non 0
                            if r.get("Lien de parenté") is not None and r.get("Lien de parenté") != "":
                                head_insuree_data["relationship_id"] = int(r.get("Lien de parenté"))
                            if r.get("Catégories_professionnelles") is not None and r.get("Catégories_professionnelles") != "":
                                head_insuree_data["professional_situation"] = professional_situations.get(
                                    int(r.get("Catégories_professionnelles"))
                                    )
                                head_insuree_data["profession_id"] = int(r.get("Catégories_professionnelles"))+1
                            if r.get("Types de formation") is not None and r.get("Types de formation") != "":
                                educations = {
                                    0: 1,
                                    1: 2
                                }
                                educ = int(r.get("Types de formation"))
                                education_id = educations.get(educ)
                                head_insuree_data["education_id"] = education_id
                            jsonext = {}
                            jsonext.update({
                                "data": {
                                    "head_insuree": head_insuree_data,
                                    "family_level": "2" if marital == "P" else "1",
                                    "location_id": current_village_id,
                                    "family_type_id": "P" if marital == "P" else "H",
                                }
                            })
                            family_data = {
                                "head_insuree": head_insuree_data,
                                "family_level": "2" if marital == "P" else "1",
                                "location_id": current_village_id,
                                "family_type_id": "P" if marital == "P" else "H",
                                "json_ext": str(jsonext)
                            }
                            head_insuree_data["json_ext"] = str(copy.copy(head_insuree_data))
                            logger.info(f"family_data {family_data}" )
                            if not parent_family:
                                # Pas de famille donc on cree direct vu que les membre menages sont en ordre
                                # c'est le premier niveau de famille ici
                                parent_family = FamilyService(self._user).create_or_update(family_data)
                                logger.info(f"family created {parent_family}")
                                familycreated += 1
                                amount_family = False
                                if r.get("Cotisationsfamilles") is not None and r.get("Cotisationsfamilles") != "":
                                    try:
                                        amount_family = Decimal(r.get("Cotisationsfamilles"))
                                    except:
                                        logger.info("Could not parse the familly contribtion amount %s",
                                            r.get("Cotisationsfamilles"))
                                contribution_plan_code = False
                                current_contribution = False
                                policy_data = {}
                                periodicity = "M"
                                value = 1
                                if amount_family == 15000:
                                    contribution_plan_code = "AMOG"
                                    value = 1
                                    periodicity = "M"
                                if amount_family == 10000:
                                    contribution_plan_code = "AMOE"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 5000:
                                    contribution_plan_code = "AMOS"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 3500:
                                    contribution_plan_code = "AMOS1"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 2500:
                                    contribution_plan_code = "AMOS2"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 2000:
                                    contribution_plan_code = "AMOS3"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 1500:
                                    contribution_plan_code = "AMOS4"
                                    periodicity = "M"
                                    value = 1
                                if amount_family == 0:
                                    contribution_plan_code = "AMS"
                                    periodicity = "Y"
                                    value = 12
                                if contribution_plan_code:
                                    current_contribution = ContributionPlan.objects.filter(
                                        code=contribution_plan_code
                                    ).first()
                                    if current_contribution and contribution_plan_code:
                                        officer = Officer.objects.filter(code=self._user._u.login_name,
                                                                        validity_to__isnull=True).first()
                                        if not officer:
                                            officer = Officer.objects.filter(validity_to__isnull=True).first()
                                        product = Product.objects.filter(id=int(current_contribution.benefit_plan_id))
                                        expiry_date = False
                                        if product:
                                            product = product.first()
                                            expiry_date = today.date() + datetimedelta(
                                                months=value
                                            )
                                            expiry_date = expiry_date - timedelta(days=1)
                                        policy_data = {
                                            "enroll_date": today.date(),
                                            "start_date": today.date(),
                                            "status": Policy.STATUS_IDLE,
                                            "contribution_plan_id": str(current_contribution.id),
                                            "value": amount_family,
                                            "audit_user_id": self._user._u.id,
                                            "product_id": current_contribution.benefit_plan_id,
                                            "periodicity": periodicity,
                                            "payment_day": 5,
                                            "officer_id": officer.id
                                        }
                                        if expiry_date:
                                            policy_data["expiry_date"] = expiry_date
                                if marital != "P":
                                    # On cree la police pour la famille si c'est pas poligame
                                    # si c'est poligame c'est la sous famille qui aura la police
                                    if current_contribution and contribution_plan_code:
                                        policy_data["family_id"] = parent_family.id
                                        logger.info("La police pour la famille %s sera cree plus tard", parent_family.id)
                                        # policy_created = PolicyService(self._user).update_or_create(policy_data, self._user)
                                        # logger.info("policy_created %s", policy_created.id)
                                        policies.append(policy_data)
                                if marital == "P":
                                    #polygamous with should create 1 subfamily
                                    jsonextsub = {}
                                    jsonextsub.update({
                                        "family_data": {
                                            "family_level": "1",
                                            "location_id": current_village_id,
                                            "family_type_id": "H",
                                            "head_insuree_id": parent_family.head_insuree.id
                                        }
                                    })
                                    sub_family_data = {
                                        "family_level": "1",
                                        "location_id": current_village_id,
                                        "family_type_id": "H",
                                        "parent_id": parent_family.id,
                                        "json_ext": str(jsonextsub),
                                        "head_insuree_id": parent_family.head_insuree.id
                                    }
                                    logger.info("Creating subfamilly for family %s", parent_family.id)
                                    sub_family = FamilyService(self._user).create_or_update(sub_family_data)
                                    familycreated += 1
                                    logger.info("sub family created %s", sub_family.id)
                                    parent_family.head_insuree.family_id = sub_family.id
                                    parent_family.head_insuree.save()
                                    if current_contribution and contribution_plan_code:
                                        policy_data["family_id"] = sub_family.id
                                        logger.info("Creation police pour la sous famille %s", sub_family.id)
                                        policies.append(policy_data)
                                        # policy_created = PolicyService(self._user).update_or_create(policy_data, self._user)
                                        # logger.info("sub family policy to create later")
                            else:
                                # On ajoute l'assurée comme membre de la famille ou sous famille existante
                                if sub_family:
                                    fid = sub_family.id
                                else:
                                    fid = parent_family.id
                                head_insuree_data["family_id"] = fid
                                head_insuree_data["json_ext"] = str(copy.copy(head_insuree_data))
                                insuree = InsureeService(self._user).create_or_update(head_insuree_data)
                                logger.info("Assuree cree %s", insuree)
                                if sub_family and number_count == 2:
                                    sub_family.head_insuree_id = insuree.id
                                    insuree.head = True
                                    insuree.save()
                                    sub_family.save()
                        logger.info("creation groupe d'itentification ok.......")
                        for police_data in policies:
                            PolicyService(self._user).update_or_create(police_data, self._user, True)
                logger.info("Fin du traitement d'import.......")
        except Exception as e:
            return InsureeImportExportService._get_general_error('FAILED TO IMPORT FILE: ', e)
        totals = {
            'sent': len(data_set),
            'created': familycreated,
        }
        errors = []
        success = True
        return success, totals, errors

class BankImportService:
    
    def __init__(self, user):
        self._user = user
        self._resource = InsureeResource(user)
    
    # Fonction pour parser les fichier BDC
    def parse_excel_bdc(self, file):
        wb = openpyxl.load_workbook(file)
        sheet = wb.active

        transactions = []
        total_kmf = Decimal("0.00")

        for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
            if not row or all(cell is None for cell in row):
                continue
            if row_idx == 0:
                continue 

            try:
                chf_id = str(int(row[0])) if isinstance(row[0], float) else str(row[0]).strip()
                date = row[2]
                description = str(row[3]).strip()
                amount_raw = row[1]

                if not amount_raw:
                    raise ValueError("Montant manquant ou vide")

                # Nettoyage du montant
                amount_str = str(amount_raw).replace(",", ".").replace(" ", "")
                amount = Decimal(amount_str)    

                transactions.append({
                    "insuree_chf_id": chf_id,
                    "date": date.isoformat() if isinstance(date, datetime) else datetime.strptime(str(date).strip(), "%m/%d/%Y").isoformat(),
                    "description": description,
                    "amount": str(amount),
                    "code_tp": "BDC",
                    "code_ext": f"bdc_{uuid4()}",
                    "code_receipt": f"receipt_{uuid4()}",
                    "label": description,
                    "fees": "0.00",
                    "amount_received": str(amount),
                    "date_payment": date.isoformat() if isinstance(date, datetime) else datetime.strptime(str(date).strip(), "%m/%d/%Y").isoformat(),
                    "payment_origin": "Banque",
                    "payer_ref": chf_id,
                })
                total_kmf += amount
            except Exception as e:
                print(f"Erreur à la ligne {row_idx}: {row} - {e}")
                continue

        return {
            "transactions": transactions,
            "total_kmf": str(total_kmf),
            "count": len(transactions),
        }

    # Fonction pour parser les fichier Exim
    def parse_excel_exim(self, file):
        wb = openpyxl.load_workbook(file)
        sheet = wb.active

        transactions = []
        total_kmf = Decimal("0.00")
        start_parsing = False

        for row in sheet.iter_rows(values_only=True):
            if not start_parsing:
                if row and any("Txn. Date" in str(cell) for cell in row if cell):
                    headers = [str(h).strip() if h else None for h in row]
                    try:
                        date_idx = headers.index("Txn. Date")
                        desc_idx = headers.index("Description")
                        credit_idx = headers.index("Credit")
                        reference_idx = headers.index("Txn.Ref No")
                    except ValueError as e:
                        raise Exception("Colonnes attendues non trouvées dans le fichier (Txn. Date, Description, Credit)") from e
                    start_parsing = True
                    continue

            if start_parsing:
                if row and str(row[1]).startswith("Opening Balance"):
                    break

                try:
                    credit_val = row[credit_idx]
                    if credit_val is not None:
                        ref = str(row[reference_idx]) if row[reference_idx] else f"ref_{uuid4()}"
                        date_str = row[date_idx]
                        if isinstance(date_str, datetime):
                            date_formatted = date_str.date().isoformat()
                        else:
                            date_formatted = datetime.strptime(date_str, "%b %d, %Y").date().isoformat()
                        transactions.append({
                            "date": date_formatted,
                            "description": str(row[desc_idx]),
                            "amount": str(Decimal(credit_val)),
                            "insuree_chf_id": ref,
                            "code_ext": ref,
                            "label": str(row[desc_idx]),
                            "code_tp": "EXIM",
                            "code_receipt": f"receipt_{uuid4()}",
                            "fees": "0.00",
                            "amount_received": str(Decimal(credit_val)),
                            "date_payment": date_formatted,
                            "payment_origin": "Banque",
                            "payer_ref": ref,
                        })
                        total_kmf += Decimal(credit_val)
                except Exception as e:
                    print(f"Erreur parsing ligne: {row} - {e}")
                    continue

        return {
            "transactions": transactions,
            "total_kmf": str(total_kmf),
            "count": len(transactions),
        }
    
    def parse_excel_other_payment(self, file):
        wb = openpyxl.load_workbook(file)
        sheet = wb.active

        transactions = []
        total_kmf = Decimal("0.00")
        start_parsing = False

        for row in sheet.iter_rows(values_only=True):
            if not start_parsing:
                if row and any("Txn. Date" in str(cell) for cell in row if cell):
                    headers = [str(h).strip() if h else None for h in row]
                    try:
                        date_idx = headers.index("Txn. Date")
                        desc_idx = headers.index("Description")
                        credit_idx = headers.index("Credit")
                        payment_source_idx = headers.index("Mode de paiement")
                        reference_idx = headers.index("Txn.Ref No")
                    except ValueError as e:
                        raise Exception("Colonnes attendues non trouvées dans le fichier (Txn. Date, Description, Credit)") from e
                    start_parsing = True
                    continue

            if start_parsing:
                if row and str(row[1]).startswith("Opening Balance"):
                    break

                try:
                    credit_val = row[credit_idx]
                    if credit_val is not None:
                        ref = str(row[reference_idx]) if row[reference_idx] else f"ref_{uuid4()}"
                        date_str = row[date_idx]
                        payment_source = str(row[payment_source_idx]).upper()
                        if isinstance(date_str, datetime):
                            date_formatted = date_str.date().isoformat()
                        else:
                            date_formatted = datetime.strptime(date_str, "%b %d, %Y").date().isoformat()
                        transactions.append({
                            "date": date_formatted,
                            "description": str(row[desc_idx]),
                            "amount": str(Decimal(credit_val)),
                            "insuree_chf_id": ref,
                            "code_ext": ref,
                            "label": str(row[desc_idx]),
                            "code_tp": payment_source,
                            "code_receipt": f"receipt_{uuid4()}",
                            "fees": "0.00",
                            "amount_received": str(Decimal(credit_val)),
                            "date_payment": date_formatted,
                            "payment_origin": payment_source,
                            "payer_ref": ref,
                        })
                        total_kmf += Decimal(credit_val)
                except Exception as e:
                    print(f"Erreur parsing ligne: {row} - {e}")
                    continue

        return {
            "transactions": transactions,
            "total_kmf": str(total_kmf),
            "count": len(transactions),
        }
    
    def parse_date(self, date_str):
        for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%b %d, %Y", "%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        raise ValueError(f"Format de date non reconnu: {date_str}")


    def log_invoice_event(self, user, invoice, event_type, message):
        event = InvoiceEvent(
            invoice=invoice,
            event_type=event_type,
            message=message,
        )
        event.save(username=user.username)

    def get_invoice_by_code(self, code):
        try:
            return Invoice.objects.get(code=code, is_deleted=False)
        except Invoice.DoesNotExist:
            return None
        
    def find_invoice(self, chf_id, get_all_invoices=False):
        try:
            insuree = Insuree.objects.get(chf_id=chf_id, validity_to__isnull=True) 
            if not get_all_invoices:
                return Invoice.objects.filter(
                    subject_id=str(insuree.id),
                    status=Invoice.Status.VALIDATED,
                    is_deleted=False,
                    thirdparty_type=73
                ).order_by("date_valid_from").first()
            # Get all invoices
            return Invoice.objects.filter(
                subject_id=str(insuree.id),
                status=Invoice.Status.VALIDATED,
                is_deleted=False,
                thirdparty_type=73
            ).order_by("date_valid_from").all()
        except Insuree.DoesNotExist:
            return None


    def valider_paiement_factures(self, factures, montant_paye):
        """
        Vérifie si `montant_paye` permet de payer les factures
        une à une dans l'ordre, sans paiement partiel.

        - factures : liste des facturees
        - montant_paye : montant total payé par l'utilisateur

        Retourne la liste des factures couvertes si OK, lève une exception sinon.
        """
        restant = montant_paye
        factures_couvertes = []

        for facture in factures:
            montant = facture.amount_total

            if restant == 0:
                # Plus rien à payer, on s'arrête
                break

            if restant < montant:
                raise Exception(
                    f"Paiement partiel non autorisé : la facture {facture.code} est de {montant} KMF "
                    f"mais le montant disponible n'est que de {restant} KMF (il manque {montant - restant} KMF)."
                )

            restant -= montant
            factures_couvertes.append(facture)

        if restant > 0:
            raise Exception(
                f"Trop perçu : après avoir couvert toutes les {len(factures_couvertes)} facture(s), "
                f"il reste un excédent de {restant} KMF."
            )

        return factures_couvertes


    def reconcile_bank_transaction(self, tx):
        chf_id = tx["insuree_chf_id"]
        return_result = []
        if not chf_id:
            raise Exception("Numéro d'assuré manquant")
        # 1. Récupération de l'assuré
        try:
            insuree = Insuree.objects.get(
                chf_id=chf_id,
                validity_to__isnull=True
            )
        except Insuree.DoesNotExist:
            raise Exception(f"Assuré introuvable avec chf_id={chf_id}")

        # 2. Récupération de la famille
        try:
            family = Family.objects.get(
                head_insuree=insuree,
                validity_to__isnull=True
            )
        except Family.DoesNotExist:
            raise Exception(f"Famille introuvable pour l'assuré {chf_id}")

        # 3. Récupération des polices actives
        policies = Policy.objects.filter(
            family=family,
            validity_to__isnull=True
        )

        # 4. Vérification suspension
        total_policies = policies.count()
        nb_suspend = policies.filter(
            status=Policy.STATUS_SUSPENDED
        ).count()

        # 5. Vérification
        if total_policies == 0:
            raise Exception("Aucune police trouvée pour cette famille.")

        if total_policies == nb_suspend:
            raise Exception(
                "La police que vous essayez de payer est suspendue."
            )

        invoice = self.find_invoice(chf_id)

        payment_date = tx.get("date")
        if not payment_date:
            raise Exception("Date de paiement manquante")
        
        try:
            if isinstance(payment_date, datetime):
                payment_date = payment_date.date()
            elif isinstance(payment_date, str):
                payment_date = self.parse_date(payment_date)
        except ValueError:
            raise Exception(f"Format de date non reconnu: {payment_date}")

        code_ext = tx.get("code_ext") or f"pay_{uuid4()}"
        code_tp = tx.get("code_tp") or "Banque"
        code_receipt = tx.get("code_receipt") or f"receipt_{uuid4()}"
        reconciliation_status = PaymentInvoice.ReconciliationStatus.RECONCILIATED
        fees = Decimal(tx.get("fees", "0.00"))
        payment_origin = tx.get("payment_origin") or "Banque"
        payer_ref = tx.get("payer_ref") or chf_id

        policy = policies.filter(
            status__in=[Policy.STATUS_IDLE, Policy.STATUS_EXPIRED, Policy.STATUS_ACTIVE]
        ).order_by("-validity_from").first()
        if not policy:
            raise Exception("Aucune police en attente, expirée ou active n'a été trouvée pour la famille dont vous souhaitez effectuer le paiement.")
        family_amount = 0
        government_amount = 0
        if policy.contribution_plan:
            if isinstance(policy.contribution_plan, (str, UUID)):
                contribution_plan_uuid = policy.contribution_plan
            else:
                contribution_plan_uuid = policy.contribution_plan.uuid
            contribution_plan = ContributionPlan.objects.filter(
                uuid=str(
                    contribution_plan_uuid
                )
            ).first()
            for calculation_rule in CALCULATION_RULES:
                # get calculation_rule amount for government
                result_signal = calculation_rule.signal_calculate_event.send(
                    sender=contribution_plan.__class__.__name__, instance=contribution_plan,
                    user=self._user, context="create",
                    family=family,
                    is_government_value=True
                )
                logger.info("result_signal %s ", result_signal)
                if result_signal[0][1]:
                    government_amount = Decimal(result_signal[0][1])
                    logger.info("government_amount %s ", government_amount)

                # get calculation_rule for familly
                result_signal = calculation_rule.signal_calculate_event.send(
                    sender=contribution_plan.__class__.__name__, instance=contribution_plan,
                    user=self._user, context="create",
                    family=family,
                    is_government_value=False
                )
                logger.info("result_signal2 %s ", result_signal)
                if result_signal[0][1]:
                    family_amount = Decimal(result_signal[0][1])
                    logger.info("family_amount = %s ", family_amount)
        payment_day = policy.payment_day
        periodicity = policy.periodicity
        logger.info("payment_day =: %s", payment_day)
        logger.info("periodicity = %s", periodicity)
        qty = 1
        if not periodicity:
            periodicity = 'M'
        if periodicity:
            if periodicity == 'Q':
                family_amount = family_amount * 3
                qty = 3
                government_amount = government_amount * 3
            elif periodicity == 'S':
                family_amount = family_amount * 6
                government_amount = government_amount * 6
                qty = 6
            elif periodicity == 'Y':
                family_amount = family_amount * 12
                government_amount = government_amount * 12
                qty = 12
        logger.info("Montant a comparer %s", family_amount)

        amount_received = Decimal(tx["amount_received"])
        if amount_received <= 0:
            raise Exception("Les montants inférieurs ou égaux à 0 ne sont pas autorisés.")
        # ---- Règle : montant non multiple -> erreur, rien n'est importé ----
        if family_amount <= 0:
            raise Exception(f"Montant de cotisation introuvable pour la police de l'assuré {chf_id}.")
        nb_periods, remainder = divmod(amount_received, family_amount)
        logger.info(" nb_periods = %s remainder = %s", nb_periods, remainder)

        # point de départ pour générer les prochaines périodes si besoin
        cursor_period_end = (
            invoice.date_valid_to if invoice else payment_date
            #payment_date c'est la date qui est sur l'import
        )
        logger.info("cursor period end = %s", cursor_period_end)
        logger.info("payment date inside import file %s", payment_date)

        all_invoices = self.find_invoice(chf_id, get_all_invoices=True)
        if invoice:
            # Vérifier les montants si au moins une facture existe, sinon on créra la facture
            results = self.valider_paiement_factures(all_invoices, amount_received)
        for _ in range(int(nb_periods)):
            invoice = self.find_invoice(chf_id)
            logger.info("Invoice found %s", invoice)
            if not invoice:
                # Plus de facture impayée -> génération d'une nouvelle période
                invoice, cursor_period_end = self._generate_next_period_invoice(
                    family, family_amount, government_amount, cursor_period_end,
                    payment_day, periodicity, payment_date
                )

            subject_type = ContentType.objects.get_for_model(invoice)
            label = tx.get("label") or f"Paiement pour {invoice.code}"
            payment_invoice = PaymentInvoice(
                code_ext=code_ext,
                code_tp=code_tp,
                code_receipt=code_receipt,
                label=label,
                reconciliation_status=reconciliation_status,
                amount_received=family_amount,
                fees=fees,
                date_payment=payment_date,
                payment_origin=payment_origin,
                payer_ref=payer_ref,
                payer_name=payer_ref
            )
            payment_invoice.save(username=self._user.username)

            detail_payment = DetailPaymentInvoice(
                payment=payment_invoice,
                subject_type=subject_type,
                subject_id=str(invoice.uuid),
                status=DetailPaymentInvoice.DetailPaymentStatus.ACCEPTED,
                fees=fees,
                amount=family_amount,
                reconcilation_id=f"recon_{uuid4()}",
                reconcilation_date=payment_date,
            )
            detail_payment.save(username=self._user.username)

            if invoice.status != Invoice.Status.RECONCILIATED:
                invoice.status = Invoice.Status.RECONCILIATED
                invoice.date_payed = payment_invoice.date_payment
                invoice.save(username=self._user.username)
            self.log_invoice_event(
                user=self._user,
                invoice=invoice,
                event_type=InvoiceEvent.EventType.PAYMENT,
                message=f"Paiement reçu pour l'assuré {chf_id}, montant {amount_received} KMF pour la date du {payment_date.strftime('%d/%m/%Y')}"
            )

            premium = self.create_premium(chf_id, payment_invoice, code_tp, invoice)
            res = {
                "invoice_code": invoice.code,
                "payment_id": payment_invoice.id,
                "detail_payment_id": detail_payment.id,
                "premium_uuid": str(premium.uuid) if premium else None,
                "status": "RECONCILIATED",
                "amount": str(invoice.amount_total),
            }
            return_result.append(res)
        return return_result


    def calculate_due_date(self, today: py_date, payment_day: int) -> py_date:
        """
        Calcule la prochaine date d'échéance en tenant compte de la période.
        
        Args:
            today: Date actuelle
            payment_day: Jour de paiement souhaité (1-31)
            period: Période en mois
            (1=mensuel, 3=trimestriel, 6=semestriel, 12=annuel)
        
        Returns:
            Prochaine date d'échéance
        """
        if payment_day < today.day:
            # Mois suivant
            if today.month == 12:
                year = today.year + 1
                month = 1
            else:
                year = today.year
                month = today.month + 1
        else:
            # Mois courant
            year = today.year
            month = today.month

        # Ajuster le jour si nécessaire
        days_in_month = calendar.monthrange(year, month)[1]
        day = min(payment_day, days_in_month)

        return py_date(year, month, day)


    def _generate_next_period_invoice(self, family, family_amount, government_amount, period_start, payment_day, period, payment_date):
        """
        Génère la facture "Cotisant" et "Etat" pour la prochaine période
        """
        period_start = (
            period_start + timedelta(days=1) if period_start else datetime.now().date()
        )
        logger.info("current period_start %s", period_start)
        date_due = self.calculate_due_date(period_start, payment_day)
        logger.info("date due : %s", date_due)
        date_due = date_due.replace(day=payment_day)
        logger.info("date due updated %s", date_due)
        periodicity = 12
        if period:
            if period == 'Q':
                periodicity = 3
            elif period == 'S':
                periodicity = 6
            elif period == 'M':
                periodicity = 1
        date_to = date_due + datetimedelta(
            months=periodicity
        )
        date_valid_to = date_to - timedelta(days=1)
        logger.info("current date_valid_to %s", date_valid_to)
        existing_invoices = False
        if family.head_insuree:
            existing_invoices = Invoice.objects.filter(
                subject_id=family.head_insuree.id,
                date_valid_from__date__gte=date_due,
                is_deleted=False
            ).exclude(status=Invoice.Status.CANCELLED)
        logger.info("existing invoices %s ", existing_invoices)
        if existing_invoices:
            return existing_invoices.first(), existing_invoices.first().date_valid_to

        base_code = f"{family.head_insuree.chf_id}_{date_due.strftime('%Y%m')}"
        timestamp = py_datetime.now().strftime('%Y%m%d%H%M%S%f')
        same_insuree_invoices = Invoice.objects.filter(
            subject_id=family.head_insuree.id
        )
        code = f"{base_code}-F-{timestamp}"
        if same_insuree_invoices:
            code = f"{code}_{same_insuree_invoices.count() + 1}"

        invoice_service = InvoiceService(user=self._user)
        family_values = {
            "code": code,
            "date_due": date_due,
            "date_valid_from": date_due,
            "date_valid_to": date_valid_to,
            "amount_net": family_amount,
            "amount_total": family_amount,
            "status": Invoice.Status.RECONCILIATED,
            "cron_job_code": code,
            "subject_id": family.head_insuree.id,
            "subject_type": "insuree",
            "thirdparty_id": family.head_insuree.id,
            "thirdparty_type": "insuree",
            "date_payed": payment_date
        }
        logger.info("family_values %s", family_values)
        result_invoice = invoice_service.create(family_values)
        if not result_invoice.get("success"):
            raise Exception(
                f"Échec de création de la facture pour la période "
                f"{date_due} - {date_valid_to}: "
                f"{result_invoice}"
            )
        invoice_line_item_service = InvoiceLineItemService(user=self._user)
        invoice_line_item_service.create(
            {
                "invoice_id": result_invoice["data"]["id"],
                "code": code,
                "ledger_account": "Cotisant",
                "quantity": periodicity,
                "unit_price": family_amount / periodicity,
                "amount_net": family_amount,
                "amount_total": family_amount,
                "cron_job_code": code,
            }
        )
        new_invoice = Invoice.objects.get(id=result_invoice["data"]["id"])
        if government_amount > 0:
            policy_holder = PolicyHolder.objects.filter(
                is_deleted=False,
                code="AFD"
            ).filter(
                Q(date_valid_to__isnull=True) |
                Q(date_valid_to__date__gte=py_datetime.today().date())
            ).first()
            logger.info("policy holder found %s", policy_holder)
            gov_code = f"{base_code}-G-{timestamp}"
            if same_insuree_invoices:
                gov_code = f"{gov_code}_{same_insuree_invoices.count() + 1}"
            gov_values = {
                "code": gov_code,
                "date_due": date_due,
                "date_valid_from": date_due,
                "date_valid_to": date_valid_to,
                "amount_net": government_amount,
                "amount_total": government_amount,
                "status": Invoice.Status.VALIDATED,
                "cron_job_code": gov_code,
                "subject_id": family.head_insuree.id,
                "subject_type": "insuree",
                "thirdparty_id": policy_holder.id,
                "thirdparty_type": "policyholder",
                "date_payed": payment_date
            }
            logger.info("gov_values %s ", gov_values)
            result_invoice = invoice_service.create(
                gov_values
            )
            logger.info("Invoice government amount created %s", result_invoice)
            item_values = {
                "invoice_id": result_invoice["data"]["id"],
                "code": gov_code,
                "ledger_account": "Etat",
                "quantity": periodicity,
                "unit_price": government_amount / periodicity,
                "amount_net": government_amount,
                "amount_total": government_amount,
                "cron_job_code": gov_code
            }
            result = invoice_line_item_service.create(
                item_values
            )
            logger.info("Invoice line gov_amount created %s", result)
        return new_invoice, date_valid_to

    def create_premium(self, chf_id, data, code_tp, invoice):
        try:
            insuree = Insuree.objects.get(chf_id=chf_id, validity_to__isnull=True)
            family = Family.objects.get(head_insuree=insuree, validity_to__isnull=True)
            policy = Policy.objects.filter(
                family=family, validity_to__isnull=True
            ).exclude(
                status__in=[Policy.STATUS_SUSPENDED, Policy.STATUS_READY]
            ).order_by("start_date").first()

            if policy:
                # Si la police n'est pas encore expirée, on la mets a jour (expiry date) et
                # on cree la premium puis on fait un return
                if policy.status == Policy.STATUS_IDLE:
                    premium_data = {
                        "audit_user_id": self._user.id,
                        "receipt": f"receipt_{uuid4()}",
                        "pay_date": data.date_payment,
                        "pay_type": "B" if code_tp in ["BDC", "EXIM", "Banque"] else "M",
                        "is_photo_fee": False,
                        "amount": data.amount_received,
                        "policy": policy
                    }
                    premium = Premium(**premium_data)
                    created = update_or_create_premium(premium, self._user)
                    logger.info(f"Contribution créée avec succès pour police: {policy.id}")
                    policy_status_premium_paid(
                        policy,
                        invoice.date_valid_from.date()
                    )
                    policy.save()
                    return created

                if policy.status == Policy.STATUS_ACTIVE:
                    # Si active, on étend la date d'expiration avec la période de grace
                    logger.info("Police %s encore active, on étend la date d'expiration", policy.id)
                    # Calcul de la date d'expiration en fonction de la périodicité
                    if policy.periodicity == Policy.MONTHLY:
                        base_expiry= invoice.date_valid_from.date() + relativedelta(months=1)
                    elif policy.periodicity == Policy.QUARTERLY:
                        base_expiry = invoice.date_valid_from.date() + relativedelta(months=3)
                    elif policy.periodicity == Policy.SEMESTER:
                        base_expiry = invoice.date_valid_from.date() + relativedelta(months=6)
                    elif policy.periodicity == Policy.YEARLY:
                        base_expiry = invoice.date_valid_from.date() + relativedelta(years=1)
                    else:
                        base_expiry = invoice.date_valid_from.date() + relativedelta(months=1)

                    product = policy.product
                    grace_days = (product.grace_period_payment or 0) * 30
                    grace_period = timedelta(days=grace_days) if grace_days else timedelta(0)
                    logger.warning("grace_period is %s ", grace_period)
                    policy.expiry_date = base_expiry + grace_period
                    logger.warning("expiry date is %s ", policy.expiry_date)
                    policy.save()
                if policy.status == Policy.STATUS_EXPIRED:
                    # Si expirée, on cree la nouvelle police, son paiement, son insuree_policy et sa date d'expiration
                    logger.info(f"Police {policy.id} expirée et période d'attente dépassée, création d'une police renouvelée.")

                    new_policy = Policy(
                        family=family,
                        product=policy.product,
                        status=Policy.STATUS_IDLE,
                        stage=Policy.STAGE_RENEWED,
                        start_date=invoice.date_valid_from.date(),
                        enroll_date=policy.enroll_date,
                        value=policy.value,
                        signature_date=policy.signature_date,
                        officer=policy.officer,
                        periodicity=policy.periodicity,
                        payment_day=policy.payment_day,
                        contribution_plan=policy.contribution_plan,
                        audit_user_id=self._user.id,
                        validity_from=TimeUtils.now()
                    )
                    new_policy.save()
                    logger.info(f"Nouvelle police renouvelée {new_policy.id} créée avec start_date={invoice.date_valid_from.date()}")

                    policy.stage = Policy.STAGE_RENEWED
                    policy.save()
                    logger.info(f"Ancienne police {policy.id} marquée comme renouvelée")

                    update_insuree_policies(new_policy, self._user.id)
                    premium_data = {
                        "audit_user_id": self._user.id,
                        "receipt": f"receipt_{uuid4()}",
                        "pay_date": data.date_payment,
                        "pay_type": "B" if code_tp in ["BDC", "EXIM", "Banque"] else "M",
                        "is_photo_fee": False,
                        "amount": data.amount_received,
                        "policy": new_policy
                    }
                    premium = Premium(**premium_data)
                    created = update_or_create_premium(premium, self._user)
                    logger.info(f"Contribution créée avec succès pour la nouvelle police: {new_policy.id}")
                    policy_status_premium_paid(
                        new_policy,
                        invoice.date_valid_from.date()
                    )
                    logger.info("Comparaison de date apres recalcul %s et la date du jour %s police traité", new_policy.expiry_date, py_datetime.today().date())
                    if new_policy.expiry_date < py_datetime.today().date():
                        # La police est quand meme expirée, il faut un autre payment pour creeer une nouvelle police
                        # afin de ne pas manquer une période non payée
                        logger.info("La police est quand meme expirée")
                        new_policy.status = Policy.STATUS_EXPIRED
                    new_policy.save()
            else:
                logger.warning(f"Aucune police active trouvée pour la famille {family.id}")
        except Exception as e:
            logger.exception(f"Erreur lors de la création de contribution pour le numéro d'assuré {chf_id}: {e}")
    
