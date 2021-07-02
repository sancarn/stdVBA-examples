# TextBoxEx for VBA UserForms

This is a text box control which uses a Microsoft Word instance as a userform control. This control offers all the features you get in word, critically **spell checking**.

![preview](docs/preview.png)

![result](docs/result.png)

## Requirements

* [stdVBA](http://github.com/sancarn/stdVBA)

## Usage

Draw a frame on your Userform. The textBox should populate the frame's location and size.

In your form code add the following:

```vb
Private textBox As uiTextBoxEx

Private Sub UserForm_Initialize()
  'Intiialise textbox
  Set textBox = uiTextBoxEx.Create(Me.TextBoxFrame)

  'You can change the text displayed via the Text property. This isn't a requirement
  textBox.Text = "Yo bruva it's ya boy Jamo and I got something really cool to tell ya " & _
                 "It's so cool that you're gunna shizzle ma nizzle isn't it?"
End Sub
```

## Roadmap

* [X] Basic text editing / display
* [ ] Disable-able (read only)
* [ ] Is there an easy way to inject and extract formatted text?

## Known issues

* Slow startup
    * Try to avoid create a huge quantity of these textboxes. A brand new word instance is created for each textbox, and that's pretty performance heavy. Instead keep 1 textbox and change the parent of it if necessary.
* Sometimes ui appears to go offscreen and there are potential sizing issues - happy for PRs to correct this.
