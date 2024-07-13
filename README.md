Tool to analyze duplicated rows in row/column sheet documents
Designed for Python 3.11

# Usage

## Initiate project

```
# Windows users
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run

Example run
```
python generate_test_data.py

python main.py --input-file test_data.xlsx --output-file output.xlsx --column-index 0 --threshold 80
```