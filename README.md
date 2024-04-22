# Toastmasters Club Assignment Manager

This Excel (.xlsm) file serves as a tool to aid the Vice President of Education in organizing member assignments for a Toastmasters club. It aims to prevent members from being assigned assignments too frequently or being accidentally overlooked, while also effectively tracking assignment history.

## Features

- **Member Assignment Management**: Members' names are recorded in the "Pathway Log" sheet and can be edited in the "Assignments" sheet.
- **Dynamic Dropdown Menu**: Clicking on an empty cell generates a dropdown menu based on the member list, simplifying the assignment process.
- **Previous Assignment Exclusion**: Members who had the same assignment in the previous meeting will not appear in the dropdown menu. This prevents repetitive assignments for members.
- **Current Assignment Exclusion**: Members who already have an assignment scheduled for the current meeting will not appear in the dropdown menus for other assignments in the same meeting, ensuring fair distribution.
- **Priority Sorting**: The dropdown menu is sorted based on the content of the last eight meetings for the same assignment. Members with fewer recent assignments are prioritized.

## Usage

1. **Maintain Member List**: Ensure the member list is up-to-date, especially when new members join the club.
2. **Assignments**: Click on the empty cell corresponding to the assignment to be scheduled.
3. **Dropdown Menu**: Arrange the assignment based on the suggested order in the dropdown menu.

## Note

Ensure macros are enabled and allow macros to run when opening the Excel file to utilize the full functionality of this tool. Refer to [Microsoft's guide](https://support.microsoft.com/zh-tw/topic/%E5%B7%B2%E5%B0%81%E9%8E%96%E6%9C%89%E6%BD%9B%E5%9C%A8%E5%8D%B1%E9%9A%AA%E7%9A%84%E5%B7%A8%E9%9B%86-0952faa0-37e7-4316-b61d-5b5ed6024216) for unblocking macros.

## License

This project is licensed under the [MIT License](LICENSE). Feel free to modify and distribute according to the terms of the license.
