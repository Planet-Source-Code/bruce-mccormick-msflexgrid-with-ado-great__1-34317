Attribute VB_Name = "Module1"
'This is my first "go" at VB. I have been programming for 30 years, but am a newbie to VB.
'It is a purchasing app I developed it for a client that owns several restaurants.
'
'This is a FlexGrid program which incorporates most of what i could find about grids on PSW,
'so if you see something that looks familiar, don't be surprised, although I have changed most
'everything to my standards. I've been "collecing" a lot of routines into the various .bas modules,
'but a lot of them have not been tested/standardized yet. I intend to convert some of them to
'classes soon . .
'
'Some of the features:
'   "Complete" navigation, including arrow keys, PgUp, PgDn, insert, delete, etc
'   Updates each row as you leave the row so you reduce the risk of losing data
'   You can arrow up/dn to records and even delete previous rows
'   Alternates row colors
'   Has row and col totals
'   Incorporates text and cbo boxes in the grid
'   What I think is a very good error routine, incl error logging to a disk file
'   All(?) ADO functions take place in the basADO module and include complete
'      reporting of the ADO error collection along with any other errors
'
'No Votes Please, but I would love comments on how I could improve my code. Feel free to use
'   anything you find of value. No credit need be given to me.
'
'Some notes:
'   On startup, the main menu appears. The only "functioning"  menu choice with this submission
'      is Purchases/Maintenance
'   The key for this file is the first 2 col (Vendor and Invoice nbr), so both must be
'      present to add/chg/del rows
'   The calendar in the grid only pops up when moving to the right so that any arrow
'      navigation doesn 't always pop it up when not needed
'  There are several invoices already in the DB that I have been testing with.
'      They are all for Vend1 and the invoice nbrs are 1, 2, 3, 4. There are others,
'      but those will get you going
'   Most of the Global Const in basMain will be going into an Admin DB as soon as I can
'      get the maint. pgm written for them.
'   It is written as a single-user, non-MDI, non-class, non-ocx, non-DLL app -
'     mostly because I too ignorant at this point to do otherwise.
'   I like to use the "Call" word on all of my calls so if I have to find all the modules that
'      are called it is easy to do with the pgm.
'
'Questions I have (that I'd love to have you answer!)
'   How do I tell for sure whether I've got a MS error or a program error ?
'      I'm using VBObjectError as a comparison on error number (basErrorMsg),
'      but I don't think that's correct.
'   Is there any way to tell through an API or ?? what the severity of an error is?
'      Should I just END after an error? I hate the idea of "Resume Next" or "GoTo 0"
'      if i don't know the severity of the error.
'   I can 't seem to align the text box in the grid so that the text in the box lines up with the
'      text in the grid. I've tried changing most of the parms and just can't get it aligned.
'      Any suggestions?
'   When I moved the app to a different folder, I lost the "connection" to the VBA functions
'      (Mid, Left, Right, Replace, etc), so I had to qualify them with "VBA". I tried looking for
'      missing references, etc, but couldn't find any. A note on Tek-Tips mentioned something
'      about changing directories, but I couldn't solve it.
'
'Thanks for checking out my code.
