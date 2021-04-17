# Codesanook.CollaborativeLocalization

A command line to create Google Sheet for creating localization file, supported language to each sheet and export to JSON

# How to use this project

## Clone and build the project
- Clone the project 
```
git clone git@github.com:codesanook/Codesanook.CollaborativeLocalization.git
```
- CD to `Codesanook.CollaborativeLocalization` folder
- Open Codesnook.CollaborativeLocalization.sln with Visual Studio
- Build the project
- You output `Codesanook.CollaborativeLocalization.exe` will be in `Codesanook.CollaborativeLocalization\Codesanook.CollaborativeLocalization\bin\Debug` folder.


## Prepare p12 file

- Downloaded p12 key file of a service account that can edit your Google sheet from https://console.developers.google.com 
- Rename it to `service-account-private-key.p12` and put it to in a working directory that are going to execute CodeSanook.CollaborativeLocalization.exe

##  Create a PowerShell script

- Start from a folder that you want to execute a PowerShell that we are goting to create 
- Given that we want to run the script inside the root of project which is `Codesanook.CollaborativeLocalization`.
- Create a `export-json-locale.ps1` and add the following script to run the exe output that we have just built.
```
# export-json-locale.ps1
$exePath = "Codesanook.CollaborativeLocalization/bin/debug/Codesanook.CollaborativeLocalization.exe"

& $exePath `
  --sheet-name "your-google-sheet-name" `
  --service-account-email "your-google-service-account" `
  --output-dir "locale" `
  --supported-languages en th
```              
- Please note that if your Google sheet does not exist it will be created automatically.
- Your current file structure:
```             
- Codesanook.CollaborativeLocalization/..
- export-json-locale.ps1 
- service-account-private-key.p12
```

## Export a JSON locale file
- Lanch a new PowerShell session and run the following command.
```
.\export-json-locale.ps1
```
- You will find an output like this:
```
> .\export-json-locale.ps1
'..\locale\en.json' exported
'..\locale\th.json' exported
> 
```
- There will be some local files in a folder that you specific with `--output-dir` parameter.


## All command line options
> .\Codesanook.CollaborativeLocalization.exe --help
```        
Codesanook.CollaborativeLocalization 1.0.0.0
Copyright c  2019

  --output-dir               (Default: .) An ouput directory of JSON local file, default to a current working directory

  --sheet-name               (Default: Codesanook.CollaborativeLocalization) Gooogle sheet name, it will be created automatically if it does exist

  --key-to-upper-case        (Default: true) Set a localization key to upper case automatically

  --supported-languages      (Default: en th) Space seperated value of supported language, these values match all sheet tabs

  --shared-to-emails         (Default: ) Space separated values emails to share a write permission to the working sheet

  --service-account-email    (Default: collaborative-localization@codesanook.iam.gserviceaccount.com) Service account email

  --application-name         (Default: Codesanook.CollaborativeLocalization) Google API application name

  --help                     Display this help screen.

  --version                  Display version information.
```   
            
