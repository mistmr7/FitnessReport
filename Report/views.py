from django.shortcuts import render, redirect
from django.views.generic import View
from .forms import LoginForm
from .fitness_report import run_program

# Create your views here.

class IndexView(View):
    template_name = 'Report/templates/index.html'
    context_data = {}


    def get(self, request, *args, **kwargs):
        user = request.session.get('User')
        self.context_data['user'] = user
        self.context_data['msg'] = ''
        if user is None:
            return redirect(to='login')
        return render(request, self.template_name, context = self.context_data)

    def post(self, request, *args, **kwargs):
        try:
            run_program()
        except PermissionError:
            user = request.session.get('User')
            self.context_data['user'] = user
            self.context_data['msg'] = 'Please close the excel file you would like to analyze and try again.'
            return render(request, self.template_name, context=self.context_data )
        return self.get(request)

class LoginView(View):
    template_name = 'Report/templates/login.html'
    context_data = {}

    def get(self, request, *args, **kwargs):
        self.context_data['form'] = LoginForm()
        return render(request, self.template_name, context=self.context_data)

    def post(self, request, *args, **kwargs):
        request.session.set_expiry(0)
        form = LoginForm(request.POST)
        if form.is_valid():
            user = form.cleaned_data['ID']
            password = form.cleaned_data['password']
            if password == 'CityOfBoulder':
                request.session['User'] = user
                return redirect(to='index')
            else:
                self.context_data['msg'] = 'Incorrect password'
                return render(request, self.template_name, self.context_data)
