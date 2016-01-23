{% extends 'setup.py.jj2' %}
{%block extras %}
if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    dependencies.append('ordereddict')
{%endblock %}

{%block additional_classifiers%}
        - 'Programming Language :: Python :: 2.6'
        - 'Programming Language :: Python :: 2.7'
{%endblock%}}

