from django.db import models

# Create your models here.


class Sgs(models.Model):
    exporterCompany = models.CharField(max_length=100, blank=True, null=True)
    exporterCompanyAddress = models.CharField(max_length=100, blank=True, null=True)
    contactPerson = models.CharField(max_length=100, blank=True, null=True)
    email = models.CharField(max_length=100, blank=True, null=True)
    phone = models.CharField(max_length=100, blank=True, null=True)
    importerCompany = models.CharField(max_length=100, blank=True, null=True)
    importCompanyAddress = models.CharField(max_length=100, blank=True, null=True)
    invoiceNoDate = models.CharField(max_length=100, blank=True, null=True)
    file = models.FileField()

    class Meta:
        verbose_name_plural = "Sgs"

    # def __str__(self):
        # return str(self.user)
