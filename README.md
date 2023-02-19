# Better-Access-Charts
Better charts for Access with charts.js

## Why Better charts for Access?

Microsoft Access badly needs modern charts. The original charts in MS Access date from the 90s of the previous century. Microsoft has given the charts in Access a lift. They called it "Modern Charts".

There are countless JavaScript libraries on the world wide web that you can use to create cool charts. This project makes use of this.
We create charts using the Chart.js library and display them in the web browser control. The whole logic is hidden in some class modules.

Take a look at the demo and let yourself be inspired by the possibilities.

![grab-landing-page](https://github.com/team-moeller/better-access-charts/blob/main/Demo.gif)

## You want to give it a try?
1. Download the [latest release](https://github.com/team-moeller/better-access-charts/releases/latest)
2. Unpack the files to a trusted folder
3. Run the database
4. Push the button: "Create Chart"

## How to integrate into your own database?
**1. Import of the class modules**

First, all modules with the name "BAC_*" must be imported from the demo database into your Access database.

**2. Insert web browser control on form**

The second step is to add a web browser control to display the chart on a form. It is best to give the control a meaningful name. This is required later in the VBA code. I like to use the name "ctlWebbrowser" for this.

The following text is entered in the "ControlSource" property: = "about: blank". This ensures that the web browser control remains empty at the beginning.

**3. First lines of code for the basic functionality**

The best thing to do is to add another button. In the click event, paste the following code:

```vba
Dim myChart As BAC_Chart  
Set myChart = BAC.Chart(Me.ctlWebbrowser)  
myChart.CreateChart  
```

* In line 1 a variable of the type BAC_Chart is declared.
* In line 2 a new instance of this class is created and the web browser control is assigned to the class module.
* The chart is created in line 3. 


When you run this code, you will see a chart with some data. At the moment no data source is assigned. In such a case, Better-Access Charts simply shows a standard data source with 6 entries. This is particularly practical for our example. We have now done a quick test and fundamentally implemented the chart.

**4. Add a data source and define the chart type**

In order for the chart to show something, it needs a [data source](https://github.com/team-moeller/better-access-charts/wiki/datasource). You can use the [DataSource.ObjectName](https://github.com/team-moeller/better-access-charts/wiki/datasource#objectname) property for this, for example. Enter the name of a table or a query that contains the data to be displayed.

You can specify one or more field names using the [DataSource.DataFieldNames](https://github.com/team-moeller/better-access-charts/wiki/datasource#datafieldnames) property. If you specify multiple field names, a data series is drawn for each field. You use the [DataSource.LabelFieldName](https://github.com/team-moeller/better-access-charts/wiki/datasource#labelfieldname) attribute to specify the field from which the names of the data points are taken.

Finally, use the [ChartType](https://github.com/team-moeller/better-access-charts/wiki/chart#charttype) property to select which of the nine possible chart types should be created.

The necessary VBA code could look like this, for example:

```vba
myChart.DataSource.ObjectName = "tbl_DemoData"
myChart.DataSource.DataFieldNames = Array("Dataset1", "Dataset2", "Dataset3")
myChart.DataSource.LabelFieldName = "DataLabel"
myChart.ChartType = chChartType.Line
```

* In line 1, the table "tbl_Demo_Data" is specified as the data source.
* Line 2 names three fields for three data series.
* Line 3 defines the name of the label field.
* In line 4, a line chart is selected as the chart type.

**5. Set further attributes for the chart**

The next step is to adapt the chart to your own needs. For example, you can define a [title](https://github.com/team-moeller/better-access-charts/wiki/title), label the [axes](https://github.com/team-moeller/better-access-charts/wiki/axis#labeltext) or adjust the [default font size](https://github.com/team-moeller/better-access-charts/wiki/font#size).

The project currently has 15 subclasses with lots of properties. You can see all of them in the [documentation](https://github.com/team-moeller/better-access-charts/wiki/documentation) on the Wiki. I have also presented the individual progress in the [blog](https://translate.google.com/translate?hl=en&sl=de&tl=en&u=https%3A%2F%2Fblog.team-moeller.de%2Fsearch%2Flabel%2FBetter%20Access%20Charts).

As you can see, there are a multitude of sources. Take a look around and make use of the options provided.
