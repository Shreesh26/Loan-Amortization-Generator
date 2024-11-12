
## **Basic Information and Running the Code**

This code was developed using Python 3.12. This documentation assumes that you have a compatible version of Python installed. 

After cloning the repository to your local system, required Python packages can be installed using the following command in the cloned repository:

```bash
pip install -r requirements.txt
```

The script file `ammortization_script.py` can then be run using the following command:

```bash
python "ammortization_script.py"
```

**Important Note**: You may use your own custom dataset, but please ensure that it follows the same format as the one found in `sample_dataset.xlsx`. In case any particular LoanTape has no mortgage renewal period, fill in the `mortgage_term_months` for that loan as `0`.

## **Output Structure**

After running the command to execute the script, the periodic, daily, and monthly amortization tables for each LoanTape in the sample dataset will be stored in the `individual amortization tables` directory. The consolidated daily and monthly amortization tables will be stored in `Consolidated Tables.xlsx` in the home directory itself.


This formatting ensures proper use of headings, code blocks, and bold text. You can copy and paste this directly into your README file.
