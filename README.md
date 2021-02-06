# Better-Access-Charts
Better charts for Access with charts.js

## Why Better charts for Access?

Microsoft Access urgently needs modern charts. The original charts in MS Access are from the 90s of the previous century. Microsoft has given the charts in Access a lift. They call it "Modern Charts".

There are many solutions for charts based on Java Script available on the web. This project makes use of this.
We create charts using the Chart.js library and display them in the web browser control. The whole logic is hidden in a class module.

Take a look at the demo and let yourself be inspired by the possibilities.

## You want to give it a try?
1. Download the [latest release](https://github.com/team-moeller/better-access-charts/releases/latest)
2. Unpack the files to a trusted folder
3. Run the database
4. Push the button: "Create Chart"

## How to integrate in your own database?
1. Import of the class module

First, the class module cls_Better_Access_Chart must be imported from the demo database into your Access database.

2. Insert web browser control on form

The second step is to add a web browser control to display the chart on a form. It is best to give the control a meaningful name. This is required later in the VBA code. I like to use the name "ctlWebbrowser" for this.


The following text is entered in the "ControlSource" property: = "about: blank". This ensures that the web browser control remains empty at the beginning.

3. First lines of code for the basic functionality

The best thing to do is to add another button. In the click event, paste the following code:

```vba
Dim myChart As cls_Better_Access_Chart  
Set myChart = New cls_Better_Access_Chart  
Set myChart.Control = Me.ctlWebbrowser  
myChart.DrawChart  
```

In line 1 a variable of the type cls_Better_Access_Chart is declared.

In line 2 a new instance of this class is created.

In line 3, the web browser control is assigned to the class module.

The chart is created in line 4. 


When you run this code, you will see a chart with some data. At the moment no data source assigned. In such a case, Better-Access Charts simply shows a standard data source with 6 entries. This is particularly practical for our example. We have now done a quick test and fundamentally implemented the chart.
