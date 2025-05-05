# OneNote-Page-Count-AddIn
I stored all the job descriptions that I applied to in OneNote, but it doesn't seem like there's any way to count # of pages in the current section, so I'm making this for that purpose. So curious to see how many jobs I've applied to, I bet it exceeded 500 :( (Update: it was 518 when I ran the tool for the first time...)


## User stories
### Counter
1. The user, in the task bar, can select from the first drop-down list [Application/Account, Notebook, Section Group, Section, Page, SubPage] from which they can count the items chosen the second drop-down list [Notebook, Section Group, Section, Page, SubPage, Word count]
2. The user can see the results in the taskbar
3. If the user's choice is invalid (i.e., the second option is not a child of the first option), prompt the user
### Numbering
4. The user can number all immediate children (by adding numbering directly to notebook content, in outline or whatever) of the current item from the drop-down list [Application/Account, Notebook, Section Group, Section, Page, SubPage]. E.g., "Number all [Page] in the current [Section]."
5. The immediate children will change automatically as the user switch option in the drop-down list
6. The user can undo their actions
