# Power Point Slide Automation  

install free python panckage python-pptx

pip install python-pptx


https://buildmedia.readthedocs.org/media/pdf/python-pptx/latest/python-pptx.pdf 


The python-pptx package provides a number of different chart types and styles for creating visualizations in PowerPoint. 

List of the available chart types in python-pptx
 
https://python-pptx.readthedocs.io/en/latest/api/enum/XlChartType.html 

It supports creating tables, adding images, shapes, and text boxes to a slide

examples - https://python-pptx.readthedocs.io/en/latest/user/charts.html 

Chart API details: 
https://python-pptx.readthedocs.io/en/latest/api/chart.html#chart-api 


In python-pptx, a Presentation object represents a PowerPoint presentation, and the slide_layouts property of a Presentation object provides access to the slide layouts available in the presentation.

Slide layouts are used as templates for creating new slides in a presentation. Each slide layout defines the positioning and styling of the different content elements on the slide, such as the title, text boxes, images, and charts.

The slide_layouts property returns a list of slide layouts available in the presentation, and the index of each layout corresponds to its position in the list.

The slide_layouts[6] refers to the 7th slide layout in the list. The slide layout at index 6 in the list is commonly known as the "Title Only" layout, as it includes a placeholder for a slide title but no placeholders for body text or other content.



XL_DATA_LABEL_POSITION
Specifies where the data label is positioned.


https://python-pptx.readthedocs.io/en/latest/api/enum/XlDataLabelPosition.html


Charts
https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-chart-overview.html


XL_CHART_TYPE
Specifies the type of a chart.

https://python-pptx.readthedocs.io/en/latest/api/enum/XlChartType.html 

Char objects
https://python-pptx.readthedocs.io/en/latest/api/chart.html


Table
https://python-pptx.readthedocs.io/en/latest/user/table.html

Table object 
https://python-pptx.readthedocs.io/en/latest/api/table.html


