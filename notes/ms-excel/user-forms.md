# User Forms

## Insertion

To insert a new User Form, right-click anywhere in the VBE Project pane, and click "Insert" > "UserForm".

![inserting a user form](/img/notes/ms-excel/user-forms/insert-userform.png)

After creating a new User Form, you should see it in the Project Pane.

![viewing a user form](/img/notes/ms-excel/user-forms/new-userform.png)


## Properties

If you want to rename the User Form, or change any of its properties, you can right-click on the User Form itself (the dotted background, not any other elements on it), then click "Properties" to reveal its properties.

![accessing user form properties](/img/notes/ms-excel/user-forms/rename-userform.png)

## Designing

Double-click on the User Form in the Project pane to enter into design mode.

In design mode, you can add resize the User Form, add controls, modify properties, etc.

![adding a label element to the form, then modifying its font properties](/img/notes/ms-excel/user-forms/userform-label-font.png)


Once you have added controls, for example a command button, you can reveal the underlying code by double-clicking on the given control in design mode.

![a completed form in design mode, with input boxes and a submission button](/img/notes/ms-excel/user-forms/userform-design-mode.png)

![code underlying the form button click event](/img/notes/ms-excel/user-forms/userform-commandbutton-click-event.png)


## Showing and Hiding

Once you have created a User Form, note its name, then call `Show` on it from somewhere else (perhaps from a workbook open event or a click event for a button on some worksheet) to programmatically launch it.


![a button on some other sheet](/img/notes/ms-excel/user-forms/userform-launch.png)

![code underling the button on some other sheet - launches the form](/img/notes/ms-excel/user-forms/userform-show.png)

To show:

```vb
UserForm1.Show
```

To hide:

```vb
UserForm1.Hide
```

<hr>

## Multi-page Objects

Insert a MultiPage object onto a User Form to create multiple layouts accessible at different times in the same space.

### Events

The MultiPage object has its own events:

name | description
--- | ---
`Enter` | Triggers when the multi-page object is first launched.
`Change` | Triggers when the page changes.

### Navigation

When navigating across different pages, set the `MultiPage.Value` as the (zero-based) index number of the page you'd like to navigate to. For example:

```vb
UserForm1.MultiPage.Value = 0 ' DISPLAY PAGE 1

UserForm1.MultiPage.Value = 1 ' DISPLAY PAGE 2

UserForm1.MultiPage.Value = 2 ' DISPLAY PAGE 3
```
