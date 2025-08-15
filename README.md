# office
Microsoft Office tools

----

### Restore Word Default Normal

````
powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command \"iwr -useb https://raw.githubusercontent.com/ulyweb/office/refs/heads/main/word/normal.ps1 | iex\"' -Verb RunAs"
````

----

### Restore Powerpoint Default Document Theme

````
powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command \"iwr -useb https://raw.githubusercontent.com/ulyweb/office/refs/heads/main/powerpoint/Default-Theme.ps1 | iex\"' -Verb RunAs"
````
