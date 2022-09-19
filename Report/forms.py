from socket import fromshare
from django import forms

class LoginForm(forms.Form):
    ID = forms.CharField(
        widget=forms.TextInput(
            attrs={
                'type': 'text',
                'placeholder': 'username',
                'class': 'form-control',
                'autocomplete': 'off',
                'autofocus': 'on'
            }
        )
    )
    password = forms.CharField(
        widget=forms.PasswordInput(attrs={"class": "form-control", "autofocus": "off", "max_length": 15})
    )