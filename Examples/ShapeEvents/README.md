# PowerPoint Shape Events

A simple example showing off the use of the WIP/ShapeHook class for PowerPoint.

# Usage

1. Open the pptm
2. Execute the `showForm` macro in dev tools
3. Click `Track Objects` button
4. Either
    * Move a shape around - this will run a macro which changes the color of the shape
    * Add a shape - this will `msgbox` the name of the newly added shape
    * Delete a shape - this will `msgbox` the name of the deleted shape
5. Click `Stop Tracking` button whenever to cease event fire.