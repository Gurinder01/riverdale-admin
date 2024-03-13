import os
from django.http import JsonResponse, HttpResponse, FileResponse
from django.shortcuts import render, redirect
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from .forms import UploadFileForm
from .utils import generate_studying_status
from django.contrib import messages

def home(request):
    form = UploadFileForm()
    return render(request, 'studyingstatus/report.html', {'form': form})

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        print('hello1')
        if form.is_valid():
            excel_file = request.FILES.get('file')
            print('hello2')
            try:
                processed_file_path = generate_studying_status(excel_file)
                processed_file_name = os.path.basename(processed_file_path)
                print('hello3')
                request.session['processed_file_name'] = processed_file_name  # Store filename in session to use in download view
                messages.success(request, 'File processed successfully.')
                return JsonResponse({'message': 'File processed successfully.', 'filename': processed_file_name})
            except Exception as e:
                messages.error(request, 'There was an error processing your file.')
                print('hello4')
                return JsonResponse({'error': str(e)}, status=500)
        else:
            messages.error(request, 'Only .xlsx files are allowed.')
            return JsonResponse({'error': 'Invalid form submission.'}, status=400)
    else:
        return JsonResponse({'error': 'Invalid request'}, status=400)


def download_file(request):
    filename = request.session.get('processed_file_name')
    if filename:
        file_path = os.path.join(settings.MEDIA_ROOT, filename)
        if os.path.exists(file_path):
            return FileResponse(open(file_path, 'rb'), as_attachment=True, filename=filename)
        else:
            messages.error(request, 'File not found.')
            return HttpResponse('File not found.', status=404)
    else:
        messages.error(request, 'No file to download.')
        return HttpResponse('No file to download.', status=400)
