function createCustomMenu(options) {

    /*
    The optons parameter should be a JSON object with the following structure
    {
        'ui' : AplicationUI,
        'title' : 'MenuTitleString',
        'menuItems' : [
            ['MenuItemString1', 'FunctionNameString1'],
            ['MenuItemString2', 'FunctionNameString2'],

            ...

            {
                'title' : 'SubMenuTitleString',
                'menuItems' : [
                    ['MenuItemString3', 'FunctionNameString3'],
                    ['MenuItemString4', 'FunctionNameString4'],
                    
                    ...
                ]
            },
            
            ...
        ]
    }
    */

    var ui = options.ui || SpreadsheetApp.getUi();
    var title = options.title.toString() || "Custom Menu";
    var menuItems = options.menuItems || [];

    var menu = ui.createMenu(title);

    // Iterrate through the menuItems and create the menu entries
    menuItems.forEach(item => {

        try {
            /*
            Checks the type of the item
            If item is an array, it will be used to add a new menuItem
            If item is an object, it will be used to create a new SubMenu
            */
            switch (Object.prototype.toString.call(item).slice(8, -1)) {
                case "Array":
                    // Check that the length is correct
                    if (item.length == 2)
                        menu.addItem(...item);
                    break;
                case "Object":
                    createSubMenu(ui, menu, item);
                    break;
            };
        } catch (err) {
            // If an error occurs, log it and continue to the next item
            Logger.log(err.toString);
        };
    });

    // Makes the completed Menu visible within the application
    menu.addToUi();

    return;
};

function createSubMenu(ui, menu, options) {
    // Extract the options for setup and itteration
    var title = options.title || "Custom Submenu";
    var menuItems = options.menuItems || ['test', 'Test'];

    // Create a new Submenu item
    var subMenu = ui.createMenu(title);
  
    // Iterrate through the menuItems and create the menu entries
    menuItems.forEach(item => {
        /*
        Checks the type of the item
        If item is an array, it will be used to add a new menuItem
        If item is an object, it will be used to create a new SubMenu within this SubMenu
        */
        switch (Object.prototype.toString.call(item).slice(8, -1)) {
            case "Array":
                // Check that the length is correct
                if (item.length == 2)
                    subMenu.addItem(...item);
                break;
            case "Object":
                createSubMenu(ui, subMenu, item);
                break;
        };
    });
    // Add the subMenu to the menu
    menu.addSubMenu(subMenu);
};
