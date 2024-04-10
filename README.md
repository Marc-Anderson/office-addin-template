# office addin

this repo exists becasue all of the addin tutorials use node with the Yeoman generator(yo office) and it doesnt make sense when you're creating something simple, it also doesnt do a good job of illustrating how addins work

## notices

this addin currently only works with outlook desktop and owa


## deployment


### prerequisites
1. web server capable of serving static html files
2. a domain and an accompanying ssl certificate


### steps to deploy
1. clone the this repo
2. update `manifest.xml`
    1. `Id`: [generate a uuid for your app](uuidgenerator.net)
    2. `ProviderName`: you or your businesses name
    3. `DisplayName`: choose a name for your addin
    4. `Description`: tell the world what the addin does
    5. update all of the urls(localhost has not been tested)
        * replace anywhere it says `https://localhost:3000` with your domain and path to the root of your web server
3. make a `.tar` of the contents of the src file only, named `office-addin.tar`
4. copy the tar to your domains web root
    * `scp C:\path\to\local\file\office-addin.tar user@ipa.dd.res.ss:/var/www/domain.com/html`
5. unpack the contents of the tar into your web root
    * `tar -xvf /var/www/domain.com/html/office-addin.tar -C /var/www/domain.com/html`
6. remove the tar
    * `rm /var/www/domain.com/html/office-addin.tar`
7. update permissions of all files and folders
    * `find /var/www/domain.com/html -type f -exec chmod 644 {} \; `
    * `find /var/www/domain.com/html -type d -exec chmod 755 {} \;`
    * the final result should look like this
        * `/var/www/domain.com/html/assets`
        * `/var/www/domain.com/html/commands`
        * `/var/www/domain.com/html/taskpane`
        * `/var/www/domain.com/html/favicon.ico`
8. side loading the addin([full details on sideloading on ms learn](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows-web))
    * outlook desktop(windows): click `file` > `manage addins` and it will open a link to your addins page
    * outlook for web(owa): go to https://aka.ms/olksideload
9. reopen outlook and your addin should be working
10. us https://aka.ms/olksideload to remove the addin when you're done


## development

### helpful addin references

anything with an * indicates they are particularly helpful

* [what is an addin](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
* [first tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/outlook-tutorial)
* [official office dev repo](https://github.com/OfficeDev)
* [*outlook tutorials on office dev repo](https://github.com/orgs/OfficeDev/repositories?type=all&q=outlook)
* [*git the gist(composing messages)](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/outlook-tutorial)
* [all about the manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview?tabs=tabid-1)
* [*manifest spec](https://learn.microsoft.com/en-us/javascript/api/manifest)
* [*how to specify support for different applications](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
* [apply shortcuts within excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts)
* [auto-open task pane](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands)


### notes about each member in the manifest

these are my(likely flawed) understanding, not the official word of msft

* Id: any unique uuid so your app can be identified
* Host>Hosts: specify which apps you want to be compatible
* Requirements>Sets>Set: defines the minimum version requirements for the office app as a whole
    * this should be the absolute bare minimum, any later versions or additional features should be defined in VersionOverrides
* FormSettings: defines entrypoint for old outlooks, ignored when using VersionOverrides
* VersionOverrides: container for elements that define host specific behaviour, or changes from the base behaviour of the addin
* VersionOverrides>Hosts>Host: specify additional apps and configurations for each
* VersionOverrides>DesktopFormFactor/MobileFormFactor/AllFormFactors(only custom functions): defines which type of device the contents works on
* this is where the fun begins, define each app and its configurations here
* VersionOverrides>[Resources](https://learn.microsoft.com/en-us/javascript/api/manifest/resources): specifies text and links that refer to files and content shared across hosts


## code snippets

display a message at the top of the email
```js
    Office.context.mailbox.item.notificationMessages.addAsync("myErrorNotification", {
        type: "errorMessage",
        message: `working to insert text sample`
    }, function(result) {
        // do nothing
    });
```


## untested: testing locally with wsl(or other web server)

in order to run on localhost you must configure ssl certs for localhost

note: when this project first began the documentation for this was terrible. i did see some additional resources online so maybe its better now

* create an npm project
* run `npm install office-addin-dev-certs`
* register **BOTH** certs in edge
    * more info on [here on stackoverflow](https://stackoverflow.com/questions/21397809/create-a-trusted-self-signed-ssl-cert-for-localhost-for-use-with-express-node)
    * these are for chrome but its essentially the same in edge
* reference **BOTH** certs in your server file
* dependencies: body-parser, cors, express, office-addin-dev-certs
* https://github.com/OfficeDev/generator-office/issues/490
* https://github.com/OfficeDev/Office-Addin-Scripts/pull/199
* https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world#configure-a-localhost-web-server-and-run-the-sample-from-localhost
