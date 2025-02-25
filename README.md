Factor Analysis Automation Tool is a Python-based tool designed to perform Exploratory Factor Analysis (EFA) and Confirmatory Factor Analysis (CFA) on survey datasets. This script automates the process of selecting optimal item groupings, computing factor loadings, and evaluating model fit indices.

📌 Features
Automates Factor Analysis: Automatically selects valid constructs from survey data based on predefined prefixes.
Performs Factor Extraction: Uses the factor_analyzer library to extract factors with varimax rotation.
Computes Model Fit Indices: Calculates key metrics such as Chi-Square, AVE (Average Variance Extracted), Composite Reliability (CR), and Construct Correlations.
Saves Results to Excel: Generates two structured Excel reports:
Factor Loadings Report: Contains factor loadings for the top-performing models.
Full Analysis Report: Includes factor loadings, model fit indices, correlation matrices, and covariance matrices.
Handles Construct Validation: Ensures each construct has a minimum number of items before inclusion in analysis.
Randomized Model Selection: Evaluates up to 1000 different valid item groupings and selects the best 50 models based on variance explained.
⚙️ How It Works
User Input:

The user provides:
The survey dataset in Excel format (.xlsx).
The number of factors to extract.
The minimum number of items per construct.
Construct Identification:

The script extracts construct prefixes from item names (e.g., IT1, IT2 → construct prefix IT).
Constructs that meet the minimum item requirement are retained for analysis.
Randomized Factor Analysis:

The script generates up to 1000 unique item groupings that satisfy the construct constraints.
Factor Analysis is performed on each item grouping using the Varimax rotation method.
Each model is evaluated for factor loadings, variance explained, and construct validity.
Model Evaluation:

The top 50 models (ranked by variance explained) are selected.
The script calculates factor correlations, reliability metrics (CR, AVE, Alpha), and model fit indices.
Saving Results:

The Factor Loadings Report is saved in Factor_Loadings.xlsx.
The Full Model Analysis Report is saved in Measurement_Model.xlsx.
📥 Installation
Install the required Python libraries before running the script:
pip install pandas numpy factor-analyzer openpyxl
📝 Usage
Place your survey dataset (.xlsx) in the working directory.
Run the script or Jupyter notebook and provide user inputs.
The results will be automatically saved as Excel files.
