REM Create a directory to sycronize repos. IE c:\GitHub

REM Use this below commented commands to get the workspace and this batch file. Then run the batch file in git bash.
git config --global user.email "26889712+Reddzman@users.noreply.github.com" REM Update with GitHub email
git config --global user.name "User Name"

REM Navigate to the base directory for syncing repos. May need to create
cd /c/GitHub

REM  Set path to the Github folder and then sync all repos including in the workspace
git clone https://github.com/UserName/RepoName
