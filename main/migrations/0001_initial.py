# Generated by Django 2.2.4 on 2021-05-08 13:55

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Sgs',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('exporterCompany', models.CharField(max_length=100)),
                ('exporterCompanyAddress', models.CharField(max_length=100)),
                ('contactPerson', models.CharField(max_length=100)),
                ('email', models.CharField(max_length=100)),
                ('phone', models.CharField(max_length=100)),
                ('importerCompany', models.CharField(max_length=100)),
                ('importCompanyAddress', models.CharField(max_length=100)),
                ('invoiceNoDate', models.CharField(max_length=100)),
            ],
            options={
                'verbose_name_plural': 'Sgs',
            },
        ),
    ]