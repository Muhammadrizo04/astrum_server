from django.shortcuts import get_object_or_404
from .models import *
from django.http import FileResponse
from django.shortcuts import render
from .model import *

def search_certificates(request):
    if 'sertificate_id' in request.GET and 'seria' in request.GET:
        sertificate_id = request.GET['sertificate_id']
        seria = request.GET['seria']

        info = None  # Переменная для хранения найденного объекта

        # Список моделей для перебора
        models_to_check = [
            ITEducator, InteriorDesign, FullStack, DataSciense, SoftWare,
            Other, NetworkAdmin, Python, PythonModule, ReactModule,
            PHPModule, CSharpModule
        ]

        # Перебираем модели и ищем объект
        for model in models_to_check:
            try:
                info = model.objects.get(sertificate_id__iexact=sertificate_id, seria__iexact=seria)
                break  # Если нашли, выходим из цикла
            except model.DoesNotExist:
                continue  # Если не нашли, продолжаем цикл

        # Рендерим результат
        if info is not None:
            return render(request, 'sertificate.html', {'info': info})
        else:
            return render(
                request,
                'error.html',
                {'sertificate_id': sertificate_id, 'seria': seria}
            )
    else:
        return render(request, 'index.html')




def detail_view_mk(request, sertificate_id):
    educator = get_object_or_404(ITEducator, sertificate_id=sertificate_id)
    return render(request, 'info_mk.html', {'educator': educator})


def detail_view_3d(request, sertificate_id):
    student_3d = get_object_or_404(InteriorDesign, sertificate_id=sertificate_id)
    return render(request, 'info_3d.html', {'student_3d': student_3d})


def detail_view_fs(request, sertificate_id):
    student_fs = get_object_or_404(FullStack, sertificate_id=sertificate_id, seria='FS')
    return render(request, 'info_fs.html', {'student_fs': student_fs})

def detail_view_bd(request, sertificate_id):
    student_bd = get_object_or_404(FullStack, sertificate_id=sertificate_id, seria='BD')
    return render(request, 'info_bd.html', {'student_bd': student_bd})

def detail_view_fd(request, sertificate_id):
    student_fd = get_object_or_404(FullStack, sertificate_id=sertificate_id, seria='FD')
    return render(request, 'info_fd.html', {'student_fd': student_fd})


def detail_view_se(request, sertificate_id):
    student_se = get_object_or_404(SoftWare, sertificate_id=sertificate_id)
    return render(request, 'info_se.html', {'student_se': student_se})


def detail_view_dt(request, sertificate_id):
    student_dt = get_object_or_404(DataSciense, sertificate_id=sertificate_id)
    return render(request, 'info_dt.html', {'student_dt': student_dt})


def detail_view_cs(request, sertificate_id):
    student_cs = get_object_or_404(CyberSecurity,sertificate_id=sertificate_id)
    return render(request, 'info_cs.html', {'student_cs': student_cs})

def detail_view_na(request, sertificate_id):
    student_na = get_object_or_404(NetworkAdmin, sertificate_id=sertificate_id)
    return render(request, 'info_na.html', {'student_na': student_na})


def detail_view_other(request, seria, sertificate_id):
    student_other = get_object_or_404(Other, seria=seria, sertificate_id=sertificate_id)
    return render(request, 'info_other.html', {'student_other': student_other})


def detail_view_python(request, sertificate_id):
    student_python = get_object_or_404(Python, sertificate_id=sertificate_id)
    return render(request, 'python.html', {'student_python': student_python})


def detail_view_modul(request, sertificate_id):
    # Список всех моделей, которые нужно проверять
    module_classes = [PythonModule, PHPModule, ReactModule, CSharpModule]
    student_module = None

    # Перебираем модели и пытаемся найти совпадение по certificate_id
    for module_class in module_classes:
        try:
            student_module = module_class.objects.get(sertificate_id=sertificate_id)
            break  # Если нашли, прекращаем поиск
        except module_class.DoesNotExist:
            continue  # Если не нашли, продолжаем проверять другие модели

    return render(request, 'module.html', {'student_module': student_module})



def download_file(request, filename):
    # Define Django project base directory
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # Define the full file path
    filepath = os.path.join(BASE_DIR, 'static', 'test', filename)

    # Set the return value of the FileResponse
    response = FileResponse(open(filepath, 'rb'))
    # Set the mime type
    response['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    # Set the HTTP header for sending to the browser
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    # Return the response value
    return response
