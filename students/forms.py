from django import forms

class SearchForm(forms.Form):
    seria = forms.CharField(max_length=3, required=False)
    sertificat_id = forms.CharField(max_length=7, required=False)
