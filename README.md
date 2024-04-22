# Toastmasters Club Assignment Manager

This Excel (.xlsm) file is designed to assist the Vice President of Education in organizing member assignments for a Toastmasters club. It aims to prevent members from being assigned assignments too frequently or accidentally overlooked, while also effectively tracking assignment history.

## Features

- Members' names are recorded in the "Pathway Log" sheet and can be edited in the "Assignments" sheet.
- Clicking on an empty cell generates a dropdown menu based on the member list.
- Members who had the same assignment in the previous meeting will not appear in the dropdown menu. For example, if Paul Lee was a Speaker in the previous meeting, he will not appear in the dropdown menu for Speaker in the current meeting.
- Members who already have an assignment scheduled for the current meeting will not appear in the dropdown menus for other assignments in the same meeting.
- The dropdown menu is sorted based on the content of the last eight meetings for the same assignment. The more frequent the occurrence, the lower the sorting priority.

## Usage

1. Maintain the member list diligently, especially when new members join.
2. Click on the empty cell corresponding to the assignment to be scheduled.
3. Arrange the assignment based on the suggested order in the dropdown menu.

## Note

Ensure macros are enabled to utilize the full functionality of this Excel file.

## License

This project is licensed under the [MIT License](LICENSE). Feel free to modify and distribute according to the terms of the license.
