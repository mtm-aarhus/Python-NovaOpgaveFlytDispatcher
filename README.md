# ReadMe

## Overview
This script automates the process of downloading, processing, and uploading Excel files between SharePoint and an orchestrator queue. It trims and cleans data before processing and ensures the correct handling of files.

## Features
- Downloads an Excel file from SharePoint.
- Reads and cleans the data.
- Prepares the data for an orchestrator queue
- Creates queue elements
- Uploads a new empty Excel file back to SharePoint.


## Requirements
- Python 3.x
- Required libraries:
  - `pandas`
  - `openpyxl`
  - `office365-rest-python-client`
  - `json`
  - `os`
  - `time`

## Installation
1. Install dependencies:
   ```sh
   pip install pandas openpyxl office365-rest-python-client
   ```
2. Ensure that necessary environment variables (`OpenOrchestratorSQL`, `OpenOrchestratorKey`) are set.


## Functions
### `sharepoint_client(username, password, sharepoint_site_url)`
- Authenticates and connects to SharePoint.

### `download_file_from_sharepoint(client, sharepoint_file_url)`
- Downloads a file from SharePoint and waits until it's fully available.

### `upload_file_to_sharepoint(client, sharepoint_file_url, local_file_path, orchestrator_connection)`
- Uploads the modified file back to SharePoint.

### `create_empty_excel(file_path)`
- Creates an empty Excel file with column headers.

# Robot-Framework V3

This repo is meant to be used as a template for robots made for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

## Quick start

1. To use this template simply use this repo as a template (see [Creating a repository from a template](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-repository-from-a-template)).
__Don't__ include all branches.

2. Go to `robot_framework/__main__.py` and choose between the linear framework or queue based framework.

3. Implement all functions in the files:
    * `robot_framework/initialize.py`
    * `robot_framework/reset.py`
    * `robot_framework/process.py`

4. Change `config.py` to your needs.

5. Fill out the dependencies in the `pyproject.toml` file with all packages needed by the robot.

6. Feel free to add more files as needed. Remember that any additional python files must
be located in the folder `robot_framework` or a subfolder of it.

When the robot is run from OpenOrchestrator the `main.py` file is run which results
in the following:
1. The working directory is changed to where `main.py` is located.
2. A virtual environment is automatically setup with the required packages.
3. The framework is called passing on all arguments needed by [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

## Requirements
Minimum python version 3.10

## Flow

This framework contains two different flows: A linear and a queue based.
You should only ever use one at a time. You choose which one by going into `robot_framework/__main__.py`
and uncommenting the framework you want. They are both disabled by default and an error will be
raised to remind you if you don't choose.

### Linear Flow

The linear framework is used when a robot is just going from A to Z without fetching jobs from an
OpenOrchestrator queue.
The flow of the linear framework is sketched up in the following illustration:

![Linear Flow diagram](Robot-Framework.svg)

### Queue Flow

The queue framework is used when the robot is doing multiple bite-sized tasks defined in an
OpenOrchestrator queue.
The flow of the queue framework is sketched up in the following illustration:

![Queue Flow diagram](Robot-Queue-Framework.svg)

## Linting and Github Actions

This template is also setup with flake8 and pylint linting in Github Actions.
This workflow will trigger whenever you push your code to Github.
The workflow is defined under `.github/workflows/Linting.yml`.

