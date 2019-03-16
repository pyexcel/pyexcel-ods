all: test

test:
	bash test.sh

format:
	isort -y $(find pyexcel_ods -name "*.py"|xargs echo) $(find tests -name "*.py"|xargs echo)
	black -l 79 pyexcel_ods
	black -l 79 tests

lint:
	bash lint.sh
