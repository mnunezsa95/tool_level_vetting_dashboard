# Tool [Level Vetting Dashboard]

# Level Validation Dashboard

## Purpose
The **Level Validation Dashboard** is an automation tool designed to determine whether a specific set of levels can be addressed under defined conditions. This tool streamlines the manual process of calculating whether these levels can be covered within a year, using only a few key inputs.

## How It Works
The Level Validation Dashboard uses JavaScript to accept a set of inputs and automatically calculate whether a specified set of levels can be covered based on the given conditions. By providing instant results, it eliminates the need for manual calculations, enhancing both efficiency and accuracy in the validation process.

## Using the Tool
Follow these steps to use the Level Validation Tool:

1. **Select the Program Name:**  
   Choose the program you wish to validate levels for from the drop-down menu.
   
2. **Select the Course:**  
   Choose the course you want to validate from the drop-down menu.
   
3. **Select the Grade:**  
   Choose the grade level from the drop-down menu.
   
4. **Enter the Levels:**  
   Input the level(s) you want to validate in the following format: `L#-L#` (e.g., L1-L3), where "L" represents an uppercase letter (A, B, C, etc.) and "#" represents a number (1, 2, 3, etc.).
   
5. **On-Ramping Option:**  
   Specify whether the grade, subject, and level combination will use "on-ramping" into specific level(s).

6. **Indicate Course Frequency:**  
   - Provide either a link (to a timetable) or enter a one- or two-digit number representing the frequency of the course for the specified level.
   - If a link is provided, the script will prompt you to specify the last column and row combination.
   - If a number is entered, the script will ask you to specify the days the course will occur using letter abbreviations (M, T, W, Th, F).
   
7. **Available Slots for the Year:**  
   - Provide either a link (to a calendar) or enter a one- or two-digit number representing the total number of available slots.
   - If a link is provided, the script will prompt you to specify the columns where the specific grade is located in each tab of the academic calendar.
   - If a number is entered, no further prompts will appear.

This tool simplifies and automates the validation process, ensuring more reliable results in less time.
