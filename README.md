# city-avg-salary
## Installation 

Clone the repository
```shell
https://github.com/open-data-kazakhstan/city-avg-salary.git
```
Requires Python 3. 

Create a virtual environment and activate it 
```bash
pip install venv
python -m venv /path/to/localrepo
```

Swicth to venv directory by using cd comand
```bash
cd /path/to/localrepo
Scripts/activate
```

Install dependecies in venv by using pip
```bash
pip install -r requirements.txt
```

Run the project:
```bash
python scripts/process.py
```

## Data 

Data collected from https://stat.gov.kz/ru/industries/labor-and-income/stat-wags/publications/57024/

We extracted the data from these sources and put it in the acrhive folder as quarters.xlsx .

We processed the original data to bring them back to normal, and extracted several aggregated datasets from them into the Data folder:

* `archive/pg16 optRK2022.xlsx` - contains information of 2022 year 
* `datapackage.json` - contains all the key information about our dataset

## Scripts

* `process.py ` - runs the script

## License

This dataset is licensed under the Open Data Commons [Public Domain and Dedication License][pddl].

[pddl]: https://www.opendatacommons.org/licenses/pddl/1-0/
