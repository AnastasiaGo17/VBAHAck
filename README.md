# VBAHAck
VBA Hack for many-to-one letter mailmerge


1. Read the Medium post for more context: https://medium.com/@ellivith/vba-hack-to-mailmerge-multiple-excel-entries-in-a-letter-ae03182cbe86

2. Read all comments in Table_to_Mail and RangetoHTML before using (comments in VBA start with ')

3. You have to have a signature saved to use signature functional

4. You have to have a permission to send messages on behalf of other mailbox to use .SentOnBehalf function

5. The macro runs up until the first empty string. Do make sure your entries "stick together".

6. Create 3 separate modules in one project for all three parts of the solution, or just copy and paste into one module
