from django.http import HttpResponse
from .models import *
from import_export import resources
from import_export.admin import ImportExportModelAdmin
from django.contrib import admin, messages
from .actions import *
from django_admin_filters import DateRange, DateRangePicker
from admin_numeric_filter.admin import RangeNumericFilter
from .model import *

class ITEducatorResource(resources.ModelResource):
    class Meta:
        model = ITEducator


@admin.register(ITEducator)
class ITEducatorAdmin(ImportExportModelAdmin):
    resource_class = ITEducatorResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'create_date', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)
    list_per_page = 50


class InteriorDesignResource(resources.ModelResource):
    class Meta:
        model = InteriorDesign


@admin.register(InteriorDesign)
class InteriorDesignAdmin(ImportExportModelAdmin):
    resource_class = InteriorDesignResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)


class FullStackResource(resources.ModelResource):
    class Meta:
        model = FullStack


@admin.register(FullStack)
class FullstackAdmin(ImportExportModelAdmin):
    resource_class = FullStackResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile, clear_student_fields]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)


class DataScienseResource(resources.ModelResource):
    class Meta:
        model = DataSciense


@admin.register(DataSciense)
class DataSciense_Admin(ImportExportModelAdmin):
    resource_class = DataScienseResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)


class SoftWareResource(resources.ModelResource):
    class Meta:
        model = SoftWare


@admin.register(SoftWare)
class Software_Admin(ImportExportModelAdmin):
    resource_class = SoftWareResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)


class NetworkAdminResource(resources.ModelResource):
    class Meta:
        model = NetworkAdmin


@admin.register(NetworkAdmin)
class NetworkAdminAdmin(ImportExportModelAdmin):
    resource_class = NetworkAdminResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile, clear_student_fields]
    list_filter = (('sertificate_id_numeric', RangeNumericFilter),)



@admin.register(Other)
class OtherAdmin(admin.ModelAdmin):
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]
    list_filter = (('create_date', DateRangePicker), ('sertificate_id_numeric', RangeNumericFilter),)

class CyberSecurityResource(resources.ModelResource):
    class Meta:
        model = CyberSecurity

@admin.register(CyberSecurity)
class CyberSecurityAdmin(ImportExportModelAdmin):
    resource_class = CyberSecurityResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile, clear_student_fields]
    list_filter = (('sertificate_id_numeric', RangeNumericFilter),)

class PythonResource(resources.ModelResource):
    class Meta:
        model = Python

@admin.register(Python)
class PythonAdmin(ImportExportModelAdmin):
    resource_class = PythonResource
    search_fields = ['ism', 'familya', 'sharif', 'sertificate_id', 'seria', ]
    list_display = ('ism', 'familya', 'sharif', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile, clear_student_fields]
    list_filter = (('sertificate_id_numeric', RangeNumericFilter),)


class PhpmoduleResource(resources.ModelResource):
    class Meta:
        model = PHPModule

@admin.register(PHPModule)
class PHPModuleAdmin(ImportExportModelAdmin):
    resource_class = PhpmoduleResource
    search_fields = ['first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id',]
    list_display = ('first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]

class PythonModuleResource(resources.ModelResource):
    class Meta:
        model = PythonModule

@admin.register(PythonModule)
class PythonModuleAdmin(ImportExportModelAdmin):
    resource_class = PythonModuleResource
    search_fields = ['first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id', ]
    list_display = ('first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]



class ReactModuleResource(resources.ModelResource):
    class Meta:
        model = ReactModule

@admin.register(ReactModule)
class ReactModuleAdmin(ImportExportModelAdmin):
    resource_class = ReactModuleResource
    search_fields = ['first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id', ]
    list_display = ('first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]


class CSharpModuleResource(resources.ModelResource):
    class Meta:
        model = CSharpModule

@admin.register(CSharpModule)
class PHPmoduleAdmin(ImportExportModelAdmin):
    resource_class = CSharpModuleResource
    search_fields = ['first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id', ]
    list_display = ('first_name', 'last_name', 'middle_name', 'seria', 'sertificate_id',)
    actions = [DownloadPptxFile]