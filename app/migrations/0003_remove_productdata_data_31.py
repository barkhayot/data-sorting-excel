# Generated by Django 4.0.6 on 2022-07-08 06:11

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('app', '0002_productdata_data_0'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='productdata',
            name='data_31',
        ),
    ]
