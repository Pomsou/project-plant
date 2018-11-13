from django.http import HttpResponse
from django.shortcuts import render


def index(request):
    context_dict = {'boldmessage': "Crunchy, creamy, cookie, candy, cupcake!"}
    return render(request, 'rango/index.html', context=context_dict)


def rango(request):
    response = "Rango says hey there partner!"
    return HttpResponse(response)


def about(request):
    response = "Documentation page"
    return HttpResponse(response)
