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
  - You can manually configure if you want the sum of different status in the "Calculos.py" script.
* Automated E-Mail services
  - Weekly or Monthly e-mail sender, containing the reports exported from the scripts.

## Planned Features

- Configuring an interface and giving more customization options for end-users with it. These options include:
  - Setting which atlassian domain and project it wants to get;
  - Configuring the account e-mail and token that will be used for authenticating the program to be executed in atlassian requests;
  - Choosing which destination it will want its files to be exported;
  - Selecting if it wants to automate the use of a weekly or monthly automated e-mail service, that will send the report for them in the time given.
  

## How To Use

To clone and run this application, you'll need [Git](https://git-scm.com) and [Python 3](https://www.python.org/downloads/). From your command line:

```bash
# Clone this repository
$ git clone https://github.com/GabrielSSGF/Report-Generator-for-Jira

# Go into the repository
$ cd Report-Generator-for-Jira

# Install dependencies
$ pip install pandas openpyxl requests

# Run the app
$ python3 Main.py
```

## Credits

This software uses the following Python libraries:

- [Pandas](https://pandas.pydata.org/) - For data analysis;
- [Openpyxl](https://pandas.pydata.org/) - For *.xlsx* manipulation;
- [Requests](https://pypi.org/project/requests/) - For API requests.
