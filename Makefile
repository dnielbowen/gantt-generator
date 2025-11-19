PY=./.venv/bin/python

generate:
	$(PY) render_gantt.py --input data/C4.xlsx --output out/gantt.html
