# RT-App

This is a simple attendance application meant to work with google apps script. 
The xlsx file can be converted to a google sheet, and RTApp.gs should be uploaded as the google apps script extension.
Student data is pasted into the 'Master Sheet'
Teacher sheets are automatically generated
Teacher sheets are regenerated daily if the "generate sheet" function is set to trigger automatically overnight; this prevents teachers from accidentally breaking the app
Teachers choose a student and a rider time, and student attendance is transferred to that teacher
An attence secretary watches the "Absent Today" sheet, to make note of who should be at school
An assistant principal or someone in charge of discipline can watch to see if anyone is marked absent, and help students get where they are supposed to be
Teachers can see live if another teacher requested one of their students, and send them along
