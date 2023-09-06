<h1 align="center">
  <br>
  <a href="https://www.atlassian.com/software/jira"><img src="https://cdn.icon-icons.com/icons2/2699/PNG/512/atlassian_jira_logo_icon_170511.png" alt="Jira" width="200"></a>
  <br>
  Report Generator for Jira
  <br>
</h1>

<h4 align="center">A simple report generator containing SLA and Status info for your  <a href="https://www.atlassian.com/software/jira" target="_blank">Jira Projects</a>.</h4>

<p align="center">
  <a href="#key-features">Key Features</a> •
  <a href="#planned-features">Planned Features</a> •
  <a href="#how-to-use">How To Use</a> •
  <a href="#credits">Credits</a>
  
</p>

## Key Features

* SLA Report
  -  Returns you a *.xlsx* file containing the percentage of SLA tickets that were broken in the time of first response and resolution in a specific week or month. It also returns the amount of tickets resolved in its respective duration;
  - The *.xlsx* file has two sheets, one with information ordered by month and the other ordered by week. Each week starts at monday.
* Status Report
  - Returns you a *.xlsx* file containing the time duration of its respective Key status and the total amount of time of all status combined;
  - The *.xlsx* file has two sheets, one with only the resolved issues and the other with resolved and unresolved issues;

## Planned Features

- Automated Email services.
  

## How To Use

To clone and run this application, you'll need [Git](https://git-scm.com) and [Python 3](https://www.python.org/downloads/). From your command line:

```bash
# Clone this repository
$ git clone https://github.com/GabrielSSGF/Report-Generator-for-Jira

# Go into the repository
$ cd Report-Generator-for-Jira

# Install dependencies
$ pip install pandas openpyxl requests

* Configure the configData.json with the necessary information

# Run the app
$ python3 SLA-Report.py
$ python3 Status-Report.py

```

## Credits

This software uses the following Python libraries:

- [Pandas](https://pandas.pydata.org/) - For data analysis;
- [Openpyxl](https://pandas.pydata.org/) - For *.xlsx* manipulation;
- [Requests](https://pypi.org/project/requests/) - For API requests.
