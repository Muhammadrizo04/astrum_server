from django.http import HttpResponse
from zipfile import ZipFile
import io

def clear_student_fields(modeladmin, request, queryset):
    for student in queryset:
#        student.frontend_qrcode = None
#        student.backend_qrcode = None
        student.pptx_file = None
        student.qr_code = None
        student.sertificate_front = None
#        student.sertificate_back = None
        student.save()

clear_student_fields.short_description = "Clear selected students' fields"


def DownloadPptxFile(modeladmin, request, queryset):
    responses = []
    
    for obj in queryset:
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        response['Content-Disposition'] = f'attachment; filename="{obj.pptx_file.name}"'

        with open(obj.pptx_file.path, 'rb') as pptx_content:
            response.write(pptx_content.read())
        
        responses.append(response)
    
    if len(responses) == 1:
        # If only one file is selected, return the response directly
        return responses[0]
    else:
        # If multiple files are selected, create a zip file and return it


        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, 'w') as zipf:
            for obj, response in zip(queryset, responses):
                filename = obj.pptx_file.name
                zipf.writestr(filename, response.content)
        
        response = HttpResponse(zip_buffer.getvalue(), content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename="pptx_files.zip"'
        return response

DownloadPptxFile.short_description = "Download PPTX Files "
