# Doing Excel Work with Python: Wrangling CSV Files

Documentation
=======================

This project is from my Advanced Analytics with Python course at Oakland University. It uses `openpyxl` to manipulate csv files
into a single excel workbook and clean the data. The steps include parsing dates, inserting dynamic excel formulas, and creating
excel formats.



Folder Structure
-----------------

Here's the folder structure that gets created by `cookiecutter-datascience-simple`:

	├── hw2_a2_wranglecsv	<- Your notebooks and scripts will live in the main project folder
		│   .gitignore					<- Common file types for git to ignore
		│   README.md					<- The top-level README for developers sing this project
		│   combinesteamlogs.py			<- This combines csv files and adds summary stats
		│   hacker_extracredit.py			<- This expands on combinesteamlogs.py to allow for multiple streams	
		│
		├───data						<- Final and intermediate data
		│   └───raw						<- The original, immutable data dump
		│		├───logs				<- This is the data for the original 
		│		└───multiplelogs		<- This is the data for the extra credit
		│
		├───docs
		│       notes.md				<- Simple markdown template for project notes
		│
		└───output
			│	readme.md				<- Guidance for using this folder
			│	BCM.xlsx				<- This is the result from the combinestreamlogs.py Folder
			│
			└───extracredit				<- These files are the result of the hacker_extracredit.py file
					BCM.xlsx
					JEF.xlsx


*This project does not use the docs folder*



Development Workflows
=======================

This project uses a cookiecutter described below. Below are the project steps provided by the cookiecutter.

Create new project
----------------------

You've already done this if you are reading this file. You ran:

```bash
cookiecutter gh:misken/cookiecutter-datascience-simple
```

Put project under version control
---------------------------------

Let's get version control set up. You don't absolutely have to do this, but you should. For the local repository, do;

```bash
git init
git add .
git commit -m "Initial commit"
```

For the remote repository, make a github repository named hw2_a1_breakeven, then do;

```bash
git remote add origin git@github.com:rkalusniak/hw2_a1_breakeven.git
git branch -M main
git push -u origin main
```

Great. Using version control is good.