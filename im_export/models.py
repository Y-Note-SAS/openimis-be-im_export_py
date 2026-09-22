from django.db import models
from django.utils import timezone as django_tz
from core.models import InteractiveUser


class BankImport(models.Model):
    """ Class Bank Import :
    Class for importation of bank extract in the system
    """
    idBankImport = models.AutoField(primary_key=True)
    importDate = models.DateTimeField(default=django_tz.now)
    user = models.ForeignKey(InteractiveUser, on_delete=models.DO_NOTHING, db_column="UserID")
    stored_file = models.FileField(upload_to="bankImports/%Y/%m/", null=True, blank=True)

    class Meta:
        db_table = "tblBankImport"
