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
