# ETSH Subsidy Report Automation
This repository stores automation code for generating the monthly FNSB Electric Thermal Storage Heater (ETSH) participant subsidy reports. The following personally identifiable information (PII) is not included, but required:
```
nna-fnsb-etsh-subsidy-reports/
└── pii/
    ├── participant-info.csv
    └── sensor-url.txt
```
With those two files, the script follows the following execution path to generate billing cycle subsidy calculations, purchase request form autofilling and appendix creation, and run a LaTeX subprocess to generate individualized reports for sending to participants. I strongly recommend building a virtual envrionment based on the package requriments in `requirements.txt` (e.g., `$ (venv) pip install -r requirements.txt`)

![program diagram](diagram.png)

Example report:

![example report](report-template.png)
