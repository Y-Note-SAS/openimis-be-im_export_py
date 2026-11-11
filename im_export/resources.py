from django.db.models import Q

from core.datetimes.ad_datetime import datetime
from im_export.apps import ImportExportConfig
from insuree.apps import InsureeConfig
from insuree.models import Insuree, Family, Gender
from insuree.uba import has_enrolment_perms
from location.models import Location

from import_export import fields, resources, widgets


class CharRequiredWidget(widgets.CharWidget):
    def clean(self, value, row=None, *args, **kwargs):
        val = super().clean(value)
        if val:
            return val
        else:
            raise ValueError('this field is required')


class ForeignkeyRequiredWidget(widgets.ForeignKeyWidget):
    def clean(self, value, row=None, *args, **kwargs):
        if value:
            return self.get_queryset(value, row, *args, **kwargs).get(**{self.field: value})
        else:
            raise ValueError(self.field + ' required')


def get_location_str_filter(str):
    return Q(code=str) | Q(name=str)  # | Q(uuid=str) | Q(id=int(str) if str.isdigit() else None)


def get_locations_ids(row, regions, districts, municipalities, villages):
    uniq_region = str(row['region'] or '').strip()
    if uniq_region not in regions:
        region_model = Location.objects.all().filter(validity_to__isnull=True) \
            .filter(get_location_str_filter(uniq_region)).first()
        if region_model:
            regions[uniq_region] = region_model.id
        else:
            raise ValueError('Location {} not found in the database'.format(uniq_region))

    district = str(row['district'] or '').strip()
    uniq_district = uniq_region + "|" + district
    if uniq_district not in districts:
        district_model = Location.objects.all() \
            .filter(validity_to__isnull=True, parent__id=regions[uniq_region]) \
            .filter(get_location_str_filter(district)).first()
        if district_model:
            districts[uniq_district] = district_model.id
        else:
            raise ValueError('Location {} not found in the database'.format(uniq_district))

    municipality = str(row['municipality'] or '').strip()
    uniq_municipality = uniq_district + "|" + municipality
    if uniq_municipality not in municipalities:
        municipality_model = Location.objects.all() \
            .filter(validity_to__isnull=True, parent__id=districts[uniq_district]) \
            .filter(get_location_str_filter(municipality)).first()
        if municipality_model:
            municipalities[uniq_municipality] = municipality_model.id
        else:
            raise ValueError('Location {} not found in the database'.format(uniq_municipality))

    village = str(row['village'] or '').strip()
    uniq_village = uniq_municipality + "|" + village
    if uniq_village not in villages:
        village_model = Location.objects.all() \
            .filter(validity_to__isnull=True, parent__id=municipalities[uniq_municipality]) \
            .filter(get_location_str_filter(village)).first()
        if village_model:
            villages[uniq_village] = village_model.id
        else:
            raise ValueError('Location {} not found in the database'.format(uniq_village))

    return regions[uniq_region], districts[uniq_district], municipalities[uniq_municipality], villages[uniq_village]


def validate_and_preprocess(dataset):
    # change the region names/codes into id.
    # a village can have homonyms therefore the parent must be taken into account
    # check if head insuree
    # check if insuree number is unique
    regions = dict()
    districts = dict()
    municipalities = dict()
    villages = dict()

    regions_ids = list()
    districts_ids = list()
    municipalities_ids = list()
    villages_ids = list()
    head = list()
    empty_indices = list()

    insuree_no_seen = set()

    for idx, row in enumerate(dataset.dict):

        row_str = ''.join([str(col or '') for col in row])
        if row_str.strip() == '':
            empty_indices.append(idx)
        else:
            # Check locations
            region_id, district_id, municipality_id, village_id = \
                get_locations_ids(row, regions, districts, municipalities, villages)
            regions_ids.append(region_id)
            districts_ids.append(district_id)
            municipalities_ids.append(municipality_id)
            villages_ids.append(village_id)

            insuree_number = row['insuree_number']
            if insuree_number not in insuree_no_seen:
                insuree_no_seen.add(insuree_number)

            if insuree_number == row['head_insuree_number']:
                head.append(True)
            else:
                head.append(False)

    for index in sorted(empty_indices, reverse=True):
        del dataset[index]

    if len(insuree_no_seen) != len(dataset['insuree_number']):
        raise ValueError('There are duplicates of insurance number in the list')

    dataset.append_col(head, 'head')
    dataset.append_col(villages_ids, 'village_id')
    dataset.append_col(municipalities_ids, 'municipality_id')
    dataset.append_col(districts_ids, 'district_id')
    dataset.append_col(regions_ids, 'region_id')


class InsureeResource(resources.ModelResource):
    insuree_headers = ['head_insuree_number', 'insuree_number', 'last_name', 'other_names', 'dob', 'sex', 'village',
                       'municipality', 'district', 'region']

    head_insuree_number = fields.Field(
        attribute='family',
        column_name='head_insuree_number',
        widget=ForeignkeyRequiredWidget(Family, field='head_insuree__chf_id'),
        readonly=True,
    )

    insuree_number = fields.Field(
        attribute='chf_id',
        column_name='insuree_number',
    )

    last_name = fields.Field(
        attribute='last_name',
        column_name='last_name',
    )

    other_names = fields.Field(
        attribute='other_names',
        column_name='other_names',
    )

    dob = fields.Field(
        attribute='dob',
        column_name='dob',
        widget=widgets.DateWidget(format=ImportExportConfig.im_export_date_format)
    )

    sex = fields.Field(
        attribute='gender',
        column_name='sex',
        widget=ForeignkeyRequiredWidget(Gender, field='code'),
    )

    village = fields.Field(
        attribute='current_village',
        column_name='village',
        widget=ForeignkeyRequiredWidget(Location, field='name'),
        readonly=True
    )

    # readonly, just for export and import validation
    municipality = fields.Field(
        attribute='current_village',
        column_name='municipality',
        widget=ForeignkeyRequiredWidget(Location, field='parent__name'),
        readonly=True
    )

    # readonly, just for export and import validation
    district = fields.Field(
        attribute='current_village',
        column_name='district',
        widget=ForeignkeyRequiredWidget(Location, field='parent__parent__name'),
        readonly=True
    )

    # readonly, just for export and import validation
    region = fields.Field(
        attribute='current_village',
        column_name='region',
        widget=ForeignkeyRequiredWidget(Location, field='parent__parent__parent__name'),
        readonly=True
    )

    def __init__(self, user, queryset=None, ):
        """
        @param user: User to be used for location rights for import and export, and for audit_user_id
        @param queryset: Queryset to use for export, Default to full quetyset
        """
        super().__init__()
        self._user = user
        self._queryset = queryset

    @classmethod
    def validate_and_sort_dataset(cls, dataset):
        validate_and_preprocess(dataset)
        return dataset.sort('head', reverse=True)

    # django-import-export 4.x calls after_init_instance / import_instance; the former
    # after_import_instance / import_obj overrides were never called (the head flag,
    # village and card_issued were not set and every row failed)
    def after_init_instance(self, instance, new, row, **kwargs):
        # for now editing is disabled
        if not new:
            raise ValueError("Insuree number already exists")

    def _check_village(self, village, label):
        """The insuree create right on `village`: inside the user's districts, globally
        or with the ENROLMENT UBA link on it (insuree.uba)."""
        if village is None or not has_enrolment_perms(
                self._user, InsureeConfig.gql_mutation_create_insurees_perms, [village]):
            raise ValueError(f"Not allowed to enrol insurees in village {label}")

    def import_instance(self, instance, row, **kwargs):
        village = Location.objects.filter(id=row['village_id'], validity_to__isnull=True).first()
        self._check_village(village, row.get('village'))
        instance.head = row['head']
        instance.current_village_id = row['village_id']

        if not instance.id:
            instance.card_issued = False
            instance.audit_user_id = self._user.i_user.id

        if not instance.head:
            # the family is looked up among the ones the user may see, and joining it is a
            # write on its village (it was any family, by its head's number)
            family = Family.get_queryset(
                Family.objects.filter(validity_to__isnull=True, head_insuree__chf_id=row['head_insuree_number']),
                self._user).select_related('location').first()
            if family is None:
                raise ValueError(f"Family of head insuree {row['head_insuree_number']} not found")
            self._check_village(family.location, row['head_insuree_number'])
            instance.family = family
            family_location = family.location
            # if the current village is the same as head's, skip the current_location
            if family_location and family_location.id == row['village_id']:
                instance.current_village_id = None

        # important to be at the end
        super().import_instance(instance, row, **kwargs)

    # django-import-export 4.x: after_save_instance(instance, row, **kwargs)
    def after_save_instance(self, instance, row, **kwargs):
        super().after_save_instance(instance, row, **kwargs)
        if instance.head:
            # without transactions a dry run must not write; with them (rolled back) the
            # family is created so that the members' rows find it (the test was inverted)
            if kwargs.get('using_transactions', False) or not kwargs.get('dry_run', False):
                instance.family = self.create_family(instance)
                instance.current_village = None
                instance.save()

    def create_family(self, instance):
        return Family.objects.create(**{
            'validity_from': datetime.now(),
            'audit_user_id': self._user.i_user.id,
            'head_insuree': instance,
            'location': instance.current_village,
            'is_offline': False,
        })

    def export(self, queryset=None, **kwargs):
        """The current insurees the user may see (Insuree.get_queryset: districts, UBA),
        not every insuree of the country."""
        if queryset is None:
            queryset = Insuree.get_queryset(self.get_queryset(), self._user)
        return super().export(queryset=queryset, **kwargs)

    def get_queryset(self):
        # unscoped on purpose: the import looks existing insurees up here (an insuree
        # out of the user's scope must still count as existing); export() scopes it
        queryset = self._queryset if self._queryset is not None else super().get_queryset()
        return queryset \
            .filter(validity_to__isnull=True) \
            .select_related('gender', 'current_village', 'family', 'family__location', 'family__location__parent',
                            'family__location__parent__parent', 'family__location__parent__parent__parent')

    class Meta:
        model = Insuree
        import_id_fields = ('insuree_number',)
        fields = ('insuree_number', 'head_insuree_number', 'last_name', 'other_names', 'dob', 'sex', 'village',
                  'municipality', 'district', 'region')
