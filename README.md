# The simplest Office.js add-in for Microsoft Excel

This is the simplest possible Office.js add-in for Microsoft Excel (task pane only). It only consists of two files (and an icon in a few sizes):

* `manifest.xml`
* `taskpane.html`

The HTML file is being served by [GitHub pages](https://docs.github.com/en/pages/quickstart). 

You can play around with the add-in by sideloading the `manifest.xml` according to the [office.js docs](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

Note that this sample uses the raw Excel JavaScript API (no Python/xlwings).

![Screenshot](/screenshot.png?raw=true)
