from django.http import HttpResponse


def index(request):
    return HttpResponse("INDEX")


def rango(request):
    response = "Rango says hey there partner!"
    return HttpResponse(response)


def about(request):
    response = "Documentation page"
    return HttpResponse(response)
