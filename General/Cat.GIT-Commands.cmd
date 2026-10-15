REM Set Git User credentials
REM If you don't want Git to expose your email you can use the Git email when configuring Git Bash
REM IE. "26889712+Reddzman@users.noreply.github.com"
git config --global user.email "Alias@domain.com"
git config --global user.name "User Name"

REM Change git bash directory
cd /c/FolderName/RepoName

REM Set up new git repo that syncs with VS Online
REM git clone {URL to Repo} {Path}
git clone https://github.com/GitAccount/RepoName

REM Depending on Client settings you may recieve SSL errors and have to turn off SSL validation
git config --global http.sslVerify false

REM If you are using Git Bash to sync on a server instead of a tool Like Visual Studio code,
REM the following commands will help sync everything
git pull
git add *
git commit -m "Add Comments Here"
git push

REM Reset a commit that fails to push to GitHub
git reset --hard HEAD~1


REM On older servers you may need to use SSH for your authentication
REM If you do it is easier to sync the repo with Git BASH
REM You will be prompted for the pass phrase created when seting up SSH
REM You will need to use the commands above to make commit and sync
eval "$(ssh-agent -s)"
ssh-add /c/GitHub/ssh/id_rsa