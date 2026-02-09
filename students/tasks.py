from celery import shared_task
from django.core.files.base import ContentFile
from .models import ITEducator
import os
from django.conf import settings
from django.core.files import File


# @shared_task
# def test_task(x, y):
#     print("TASK WORKED", x + y)
#     return x + y

@shared_task
def generate_student_files(student_id):
    student = ITEducator.objects.get(id=student_id)

    if student.pptx_file and student.sertificate_front and student.sertificate_back:
        return


    if not student.qr_code:
        student.generate_qr_code()

    if not student.sertificate_front:
        background_image_path_front = os.path.join(settings.MEDIA_ROOT, 'template/malaka-1.png')
        qr_code_image_path_front = os.path.join(settings.MEDIA_ROOT,
                                                f'malaka_qrcode/qr_code-{student.sertificate_id}.png')
        output_image_path_front = os.path.join(settings.MEDIA_ROOT,
                                            f'malaka_ser_front/certificate-{student.sertificate_id}.png')
        student.overlay_qr_code_front(background_image_path_front, qr_code_image_path_front, output_image_path_front,
                                (115, 1435), 390)
        # student.sertificate_front = os.path.relpath(output_image_path_front, settings.MEDIA_ROOT)
        with open(output_image_path_front, "rb") as f:
            student.sertificate_front.save(
                f"certificate-{student.sertificate_id}-front.png",
                File(f),
                save=False
            )


    if not student.sertificate_back:
        background_image_path_back = os.path.join(settings.MEDIA_ROOT, 'template/malaka-2.png')
        qr_code_image_path_back = os.path.join(settings.MEDIA_ROOT,
                                            f'malaka_qrcode/qr_code-{student.sertificate_id}.png')
        output_image_path_back = os.path.join(settings.MEDIA_ROOT,
                                            f'malaka_ser_back/certificate-{student.sertificate_id}.png')
        student.overlay_qr_code_back(background_image_path_back, qr_code_image_path_back, output_image_path_back,
                                (2535, 1445), 390)
        student.sertificate_back = os.path.relpath(output_image_path_back, settings.MEDIA_ROOT)

        # student.sertificate_back = os.path.relpath(output_image_path_back, settings.MEDIA_ROOT)
        with open(output_image_path_back, "rb") as f:
            student.sertificate_back.save(
                f"certificate-{student.sertificate_id}-back.png",
                File(f),
                save=False
            )


    if not student.pptx_file:
        pptx_buffer = student.generate_certificate()
        student.pptx_file.save(f'{student.seria}-{student.sertificate_id}-{student.familya}-{student.ism}-{student.sharif}.pptx',
                            ContentFile(pptx_buffer.read()), save=False)
    student.save()
