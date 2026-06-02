# PackIt
> A packing list generator that creates a downloadable spreadsheet with clothing, toiletries, and liquids based on temperature data and trip duration.

![Live App Link](https://pack-it.streamlit.app/)

### Overview
Everyone wants to go on vacation. But most don't look forward to packing their belongings without bringing too much or forgetting something. PackIt takes the stress out of trip preparation by providing a list of liquids, toiletries, and clothing the user should take based on their travel destination, trip duration, and month of travel. All data is pulled from a 2024 database of average monthly temperatures for each U.S. state. When generated, this packing list can be easily downloaded as a spreadsheet for those who like to print it and physically check things off. 

### Key Features
- **Clean UI:** Hosted by Streamlit, PackIt has a very user-friendly design with dropdowns for each form item.
- **Customization:** Even though PackIt handles most of the decision making, it still gives users the option to select the specific toiletries and liquids they will need for their trip.
- **Downloadable Packing List:** Upon generation of the packing list, the user can download an .xlsx document that is formatted to be printed off and make packing a breeze.

### Languages and Libraries
- **Python:** Allowed me to strengthen my Python skills after taking an introductory programming course.
- **Streamlit:** Made creating a web application with Python quick and efficient.
- **openpyxl:** Essential for creating and editing .xlsx files in Python.

### Future Developments
- **Updated Dataset:** Increase program accuracy by fetching data from an API.
- **More Variety:** Add more generation options for trip durations and temperature ranges.
- **Personal Item Section:** Include a section for things users need in their backpack or purse.