Sub Main
	'===================================================
	'Move node(s) to location without the X Y Z confirmation box.
	'
	'Written by AMJ 4.22.2020
	'===================================================
    Dim App as femap.model
    Set App = feFemap()

    Dim nodeSet As femap.Set
    Set nodeSet = App.feSet

    Dim iterNode As femap.Node
    Set iterNode = App.feNode

    Dim rc As zReturnCode

    lblStart:
	rc = nodeSet.Select(FT_NODE,True,"Select node(s) to move")

	If rc = FE_CANCEL Then
		Exit Sub
	End If

	Dim xyz As Variant

	rc = App.feCoordPick("Move node(s) to this location",xyz)

	If rc = FE_CANCEL Then
		Exit Sub
	End If

	nodeSet.Reset
	While nodeSet.Next
		iterNode.Get(nodeSet.CurrentID)
		iterNode.xyz = xyz
		iterNode.Put(iterNode.ID)
	Wend

	App.feViewRegenerate(0)

	GoTo lblStart


    
End Sub
