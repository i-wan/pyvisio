init:
	pip install -r requirements.txt

test:
	py.test tests

clean:
	rm -rf ./dist ./PyVisio.egg-info 
