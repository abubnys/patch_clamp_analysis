# Patch Clamp Analysis
matlab scripts for analyzing and plotting patch clamp data

### Overview
This script takes the electrophysiology data from whole cell patch clamp experiments that has been saved to multiple worksheets of an excel spreadsheet file and pretty plots the results according to the type of experiment that was performed. There are 4 kinds of experiments possible:
1. epsc: the spontaneous membrane voltage of the neuron was recorded for 60 seconds, no current or voltage was injected
2. I-clamp: increasing steps of current were injected into the neuron to elicit action potentials
3. IV plot: increasing voltage steps were injected into the neuron to elicit changes in current, the magnitude of the current response to voltage step is used to generate an I-V plot of voltage-gated sodium currents, voltage-gated fast potassium currents, and voltage-gated slow potassium currents.
4. ntx: the spontaneous membrane voltage of the neuron was recorded in response to the application of various neurmodulatory agents, informatiton about the timing and type of drug application is contained in the labels.xlsx file in this repository.

### Running the script
Before running the script, specify the location of the excel spreadsheets containing the data and drug injection labels (if applicable) as the variable `path`. Once initialized, the script will prompt the user to specify which worksheet they want to analyze, and the worksheet that the drug injection labels are located in (if applicable). 
```
page of worksheet? 
```
If the identified experimental type is epsc, the program will plot the spontaneous activity and then prompt the user to alter the y range if necessary
![initial epsc plot] (/readme_screenshots/epsc1.png)

