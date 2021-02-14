# SecureADODB Fork
I started working on SecureADODB to
 - practice VBA OOP, extending both my understanding and practicall skill
 - deeply learn the original SecureADODB project
 - to see if I can identify improvement opportunities and implement them.

So, this project is largely a learning exercise for me, but I also plan to use it in another project, which can benefit from an ADODB VBA library.  
 
 ## Class diagram and mapping to ADODB

The class diagram below shows the core SecureADODB classes (this fork, blue) and the mapping to the core ADODB classes (green).

![SecureADODB-ADODB](https://github.com/pchemguy/RDVBA-examples/blob/develop/UML%20Class%20Diagrams/SecureADODB%20-%20ADODB%20Class%20Mapping.svg)

DbManager class shown at the bottom is functionally similar to the UnitOfWork class from the base project, while DbRecordset class has been added in this fork.
