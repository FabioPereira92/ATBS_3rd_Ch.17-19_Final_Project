# Auto Report Generator
A program that reads a .csv file with data to be analysed, a .json file with configuration information and a .xml file with settings and generates a .docx report based on the information from those files. It also schedules the periodic creation of the report and logs an entry in a log .csv file for every report generated.

## Features
- Reading and parsing a .csv, a .json and a .xml files
- Analysing and summarizing the data from the files
- Generating a .docx report
- Scheduling the periodic generation of reports
- Generating and appending to a log .csv file for every report generation
  
## How to run
1. Clone this repository or download the file `autoReportGenerator.py`.
2. Run in terminal.

## Example Usage
### Pre-run folder structure
![pre-run folder structure1](https://github.com/user-attachments/assets/685c8147-3bcb-43ce-99dc-1ead15c0ee57)
![pre-run folder structure2](https://github.com/user-attachments/assets/93ae3992-8cf6-467b-abf9-62877665282a)

### Content of the csv file 
![content of the csv file](https://github.com/user-attachments/assets/c9160ddb-55e8-4f4e-b8ac-4c1858430ed8)

### Content of the json file 
![content of the json file](https://github.com/user-attachments/assets/df370cfd-bd44-4439-87d6-6dd961aa0333)

### Content of the xml file 
![content of the xml file](https://github.com/user-attachments/assets/c1a32aca-e170-4799-8fa5-9cfe5c9eb2f4)

### post-run folder structure
![post-run folder structure 1](https://github.com/user-attachments/assets/cce21f91-f96f-41d3-bd18-22f03314ad62)
![post-run folder structure 2](https://github.com/user-attachments/assets/f6230f76-4434-4cd6-a71c-ccb0cddcf4d7)
![post-run folder structure 3](https://github.com/user-attachments/assets/303b2ee3-d0cd-4c45-9be2-2acd669e1d64)
![post-run folder structure 4](https://github.com/user-attachments/assets/5abe3f9a-df22-4b85-a51b-8932e52a44e8)
![post-run folder structure 5](https://github.com/user-attachments/assets/de15c49e-daea-43fc-a17e-991ddef00654)

### Content of the report
![content of the report pg1](https://github.com/user-attachments/assets/45fd28d4-7969-4d78-b987-7bc59ec7a799)
![content of the report pg2](https://github.com/user-attachments/assets/ab357b4a-deb1-4677-8e02-27003c865830)

### Content of the log file
![content of the log file](https://github.com/user-attachments/assets/e2665a06-caeb-482b-8c92-5b5eb97ea17d)

## Tech Stack
- Python 3.13
- Standard library: os, csv, json, datetime, time, pathlib
- Third-party libraries: xmltodict, python-docx

## License
MIT
