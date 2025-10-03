# bible-verse-slide-generator

<b>Project Purpose:</b> Streamline the process of copy-paste of Bible Verses to PowerPoint Slides into slides automation using Python script.

<b>Project Working Mechanism:</b> 
1) Download bible in CSV file format
2) Read the bible as a csv file using pandas df
3) Extract the required bible verses into another CSV file
4) Write the bible verses into presentation slides using `python-pptx` library.

> [!IMPORTANT]
> Does it mean we don't need to read bible when preparing anymore? NO. This is just an efficiency tool. <b>It is your responsibility to read the word of God.</b>

------------------------------------------------------------

# Installation Guide

## Python portion

1) Open Command Prompt on Windows:

	Press Windows Key + R, type cmd, press Enter

	Or search “Command Prompt” in the Start Menu

------------------------------------------------------------

2) Check Python is Installed

	In the command window, type:
		
		python --version
	or:
		python3 --version

	✅ If it shows something like Python 3.11.5, you’re good.

	❌ If it says “Python not found”, you’ll need to install Python first.

------------------------------------------------------------

3) Create a Virtual Environment & Install Required Packages

	<pre>python -m venv .venv 

	pip install -r requirements.txt</pre>

	If that doesn’t work, try:

		python3 -m venv .venv 

		pip3 install -r requirements.txt

------------------------------------------------------------

4) Activate the Virtual Environment 

	<b>(For Windows)</b> <pre>./.venv/Scripts/activate</pre>


------------------------------------------------------------

## Bible Portion

1) Go to [Bible SuperSearch](https://www.biblesupersearch.com/bible-downloads/) and download your preferred version in CSV format.

2) The CSV file might contain some symbols that are considered noises. Make changes to `clean_bible.py` and run it.

3) After the bible (in CSV format) is clean, open the below script.
	
	i) Make changes to `ref` in the script by adding your required verses.
	
	ii) Run the script below to generate a csv file containing the verses. 
	<pre>python3 api_gen_csv.py</pre>

4) Depending on your need, run either `gen_bible_verse.py` or `gen_main_ppt.py` to generate the powerpoint slides using the verses generated from (3).
> [!WARNING]  
> I added English verses csv externally for both scripts, feel free to remove the df with `bible_verses_en.csv`. 

