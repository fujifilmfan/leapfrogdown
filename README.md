leapfrogdown
=============

A VBA script for Excel that allows an arbitrary number of rows to be shifted downward to the next available blank row while 'leap frogging' over 'fixed' content.

MOTIVATION
I use Excel workbooks to create a schedule of my daily lesson plans for the whole year.  One column of the schedule contains dates.  Other columns include things like warm-up activities, lesson content, homework, etc.  If I get ahead or behind by a period, I want to be able to shift future lessons up or down accordingly.  Each week is in a block, so simply inserting or deleting a row does not work.  In addtion, some items on the schedule (such as holidays, teacher in-service days, etc.) must remain fixed.  Copying and pasting dozens of periods at a time is tedious, and maintaining the variety of border formats and conditional formatting in the spreadsheet complicates things further.

SOLUTION: Overview
The user chooses the span of columns to be copied and pasted (in the code) and which rows are allowed to move and which are not (in the spreadsheet).  The user also chooses the insertion row (in the spreadsheet), the point at which the user wants a blank row to appear.  The script then searches for the first available blank row (by checking a particular column designated by the user in the code), finds the first row above it that is allowed to move, and copies and pastes it into the blank row.  The process repeats until the designated insertion row has been moved.
