# Make-Import-ExchangeCert
Powershell script for importing Certbot Exchange certificate.
############################################################################
#         Microsoft Exchange Certbot Certificate Replacement Script        #
#                 Written in December of 2022 by Jason Beaver              #
#                                                                          #
#  Prerequisites:  OpenSSL for Windows (x64), Microsoft Exchange, Certbot  #
#                     Operating System:  Windows Server                    #
############################################################################################################################
#  Insert into Task Scheduler to run once per week.                                                                        #
#  Suggested Parameters:                                                                                                   #
#  -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "C:\Windows\System32\Scripts\Make-Import-ExchangeCert.ps1" #
#                                                                                                                          #
#  Make a directory called "pfx" under C:\Certbot\live\<your domain here>\                                                 #
#  The user this runs as should have full permissions to that directory.                                                   #
#  MAKE SURE TO HAVE CERTBOT GENERATE AN RSA CERTIFICATE - EXCHANGE HATES OTHER TYPES!                                     #
############################################################################################################################
