PY=./.venv/bin/python

FILE="../pulsevpn/config/Downloads/Component 4- Ground Tools and User Interface.xlsx"
FILE="./data/Component 4- Ground Tools and User Interface (1).xlsx"
FILE="./data/Phase 2- Component 4- Ground Tools and User Interface.xlsx"

generate:
	$(PY) render_gantt.py --input $(FILE) \
		--exclude-bucket "*onitoring*" \
		--output out/gantt4.html
