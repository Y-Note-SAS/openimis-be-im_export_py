"""
Security parity of the insuree import / export (REST).

- The export dumped every current insuree of the country: it is narrowed to the insurees
  the user may see (Insuree.get_queryset: districts, ENROLMENT UBA).
- The import resolved villages with no scope and attached members to any family by its
  head's number: each row needs the insuree create right on its village and on the
  family's (insuree.uba.has_enrolment_perms), the family is looked up through
  Family.get_queryset.
- With django-import-export 4.x the former import_obj / after_import_instance hooks were
  never called (every row failed): they run as import_instance / after_init_instance.

Fixture: core.test_helpers.create_security_tree -> R1>D1>W1>V1 / R2>D2>W2>V2, a family in
each village; `user_a` holds the insuree rights globally and is assigned district D1 only.
"""
from django.conf import settings
from django.core.cache import cache
from django.test import override_settings
from rest_framework.test import APITestCase
from tablib import Dataset

from core.test_helpers import create_security_tree, create_right_only_user, get_test_token, perm_ids
from im_export.resources import InsureeResource
from insuree.apps import InsureeConfig
from insuree.models import Insuree
from insuree.test_helpers import create_test_insuree

RIGHTS = list(dict.fromkeys(perm_ids(InsureeConfig.gql_query_insurees_perms)
                            + perm_ids(InsureeConfig.gql_mutation_create_insurees_perms)))


def _row(head, number, village):
    ward = village.parent
    district = ward.parent
    return (head, number, f"Last{number}", "Other", "01/02/2000", "M",
            village.name, ward.name, district.name, district.parent.name)


@override_settings(ROW_SECURITY=True)
class InsureeImportExportScopeTest(APITestCase):

    @classmethod
    def setUpTestData(cls):
        cls.t = create_security_tree("IEX")
        for key, village in (("a", cls.t.v1), ("b", cls.t.v2)):
            insuree = create_test_insuree(
                with_family=True, is_head=True,
                custom_props={"chf_id": f"IEX0000{1 if key == 'a' else 2}", "current_village": village},
                family_custom_props={"location": village})
            setattr(cls, f"ins_{key}", insuree)
        cls.user_a = create_right_only_user("iexuser", RIGHTS, [cls.t.d1.id])

    def setUp(self):
        cache.clear()

    def _import(self, *rows, dry_run=False):
        resource = InsureeResource(self.user_a)
        dataset = resource.validate_and_sort_dataset(
            Dataset(*rows, headers=InsureeResource.insuree_headers))
        return resource.import_data(dataset, dry_run=dry_run, use_transactions=True)

    @staticmethod
    def _errors(result):
        errors = [str(error.error) for _index, row_errors in result.row_errors() for error in row_errors]
        errors += [str(row.error) for row in result.invalid_rows]
        return " | ".join(errors)

    def test_export_within_the_districts(self):
        numbers = {row["insuree_number"] for row in InsureeResource(self.user_a).export().dict}
        self.assertIn("IEX00001", numbers)
        self.assertNotIn("IEX00002", numbers)

    def test_rest_export_within_the_districts(self):
        response = self.client.get(
            f"/{settings.SITE_ROOT()}im_export/exports/insurees?file_format=csv",
            HTTP_AUTHORIZATION=f"Bearer {get_test_token(self.user_a)}")
        self.assertEqual(200, response.status_code, response.content)
        content = response.content.decode()
        self.assertIn("IEX00001", content)
        self.assertNotIn("IEX00002", content)

    def test_import_in_scope(self):
        result = self._import(_row("IEX10001", "IEX10001", self.t.v1),
                              _row("IEX10001", "IEX10002", self.t.v1))
        self.assertFalse(result.has_errors() or result.has_validation_errors(), self._errors(result))
        head = Insuree.objects.get(chf_id="IEX10001", validity_to__isnull=True)
        member = Insuree.objects.get(chf_id="IEX10002", validity_to__isnull=True)
        self.assertTrue(head.head)
        self.assertEqual(self.t.v1.id, head.family.location_id)
        self.assertEqual(head.family_id, member.family_id)
        self.assertFalse(member.card_issued)

    def test_import_dry_run_with_members(self):
        result = self._import(_row("IEX20001", "IEX20001", self.t.v1),
                              _row("IEX20001", "IEX20002", self.t.v1), dry_run=True)
        self.assertFalse(result.has_errors() or result.has_validation_errors(), self._errors(result))
        self.assertFalse(Insuree.objects.filter(chf_id__in=["IEX20001", "IEX20002"]).exists())

    def test_import_in_a_village_out_of_scope_refused(self):
        result = self._import(_row("IEX30001", "IEX30001", self.t.v2))
        self.assertTrue(result.has_errors(), "an out-of-scope village was accepted")
        self.assertIn("not allowed", self._errors(result).lower())
        self.assertFalse(Insuree.objects.filter(chf_id="IEX30001").exists())

    def test_join_a_family_out_of_scope_refused(self):
        # a member in the user's village, attached to the family of a head out of scope
        result = self._import(_row("IEX00002", "IEX40001", self.t.v1))
        self.assertTrue(result.has_errors(), "a family out of scope was joined")
        self.assertIn("not found", self._errors(result).lower())
        self.assertFalse(Insuree.objects.filter(chf_id="IEX40001").exists())
        self.assertEqual(1, Insuree.objects.filter(family=self.ins_b.family, validity_to__isnull=True).count())
