# A helper to remove a notification to accept the cookies
The table contains VBA code that uses “SeleniumBasic” to open the desired web page with Chrome and:
- Write out all cookies that were set when you press an “Accept cookies” button (use “Write out all cookies” button)
- Check if the selected cookies remove the “Accept cookies” popup window after reloading of the page (use “Check selected cookies” button)

In most cases the cookie named “consentUUID” helps to remove annoying “Accept cookies” popup window. The parameters are obtained by pressing of “Write out all cookies” button: 
- After pressing of “Write out all cookies” button the page will be shown and 
- You have to press the “Accept cookies” button. 
- Afterwards you have to press “Ok” at the VBA message box “Accept cookies ant press Ok”.
- In columns A, B and C the names, values and domains of all cookies that were set by pressing on will appear.
- You have to select from the list of all cookies the necessary cookies that will remove the “Accept cookies” popup window.
- In most cases the cookie with a name “consentUUID” removes it. 

If you have selected one or several cookies, you can check if they remove the “Accept cookies” popup window by pressing of “Check selected cookies” button:
- The list of cookies to check should start at line 3 and contain no lines without data between them.
- The list should have the same structure as the original list of all cookies: name, value and domain.
- After pressing of “Check selected cookies” button the cookies from the list will be added and the web page will be reloaded to show if the adding of the cookies removed the “Accept cookies” popup window.
- After page is reloaded, the VBA message box “Check the result!” appears and you have to check the result of adding of the cookies. By pressing “Ok” the program finishes and the Chrome disappears. 
