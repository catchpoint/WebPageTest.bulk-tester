# WebPageTest Bulk Tester

Use Google Docs to test multiple URLs using WebPageTest (either webpagetest.org if you have an API key, or another publicly accessible instance)

Each test uses one of a defined set of parameters (a scenario) so tests can either share the same parameters, or use different sets depending on need

If a test completes successfully, selected values from the results are extracted and added to the Tests tab.

Comments, suggestions, improvements etc. welcome

There are brief instructions below but for more detailed one see the relevant Performance Advent Calender post -  http://calendar.perfplanet.com/2014/driving-webpagetest-from-a-google-docs-spreadsheet/


# Using

1. Make a copy of Spreadsheet

	https://docs.google.com/spreadsheets/d/1iAmQLoazddTPu08NHzp2Zk2sXzpprhIdLCpTEt3gVYU/

	The spreadsheet is shared read-only so you'll first need to make a copy

2. Configuring Spreadsheet - Settings Tab

	Add your own WPT API key
	Customise the parameters and results maps to include the parameters you want to specify and the results values to be extracted

3. Defining Tests - Scenarios Tab

	Create one or more test scenarios (a scenario is a named set of test parameters)
	First column must always be the name of the scenario, other columns are defined by the Parameters map in the Settings tab

4. Specifying URLs to be Tested - Tests Tab

	Add URLs to be tested in the first column, and scenario in the second (a drop down can be created via the Data > Validation menu, or just copy cell from previous row)

5. Running Tests

	Once the URLs to be tested and the corresponding scenario have been defined, choose 'Run Tests' from WebPageTest menu (on first run the app will need to be authorised and the test re-submitted)

	Once the tests have been submitted the results will be polled once a minute until they have all completed 

	To re-run a test delete the WPT URL and then choose 'Run Tests' from WebPageTest menu
	To re-retrieve the results delete the status (and corresponding results) and choose 'Get Results' from the WebPageTest menu

