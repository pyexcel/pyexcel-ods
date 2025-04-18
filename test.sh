#/bin/bash
pip freeze
pytest --verbosity=3 --cov=pyexcel_ods --doctest-glob=*.rst
