# 6001CMD Machine Learning: Hotel Booking Demand Analysis

Individual coursework (CW1) for 6001CMD Machine Learning.

## Dataset
Hotel Booking Demand (Antonio, de Almeida & Nunes, 2019), obtained via the TidyTuesday repository:
https://raw.githubusercontent.com/rfordatascience/tidytuesday/master/data/2020/2020-02-11/hotels.csv

## Structure
- `data/raw/`: original dataset (unmodified)
- `data/processed/`: outputs of the preprocessing pipeline
- `notebooks/`: analysis notebooks for Tasks 1 to 4
- `figures/`, `tables/`: exported evidence used in the report
- `src/`: reusable helper functions

## Reproducing the analysis
1. `python -m venv .venv`
2. Activate the environment and run `pip install -r requirements.txt`
3. Open the notebooks in `notebooks/` in order using the `.venv` kernel
