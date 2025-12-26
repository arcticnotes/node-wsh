Function Dict2VBArray( dict)
	Dim vbArray()
	ReDim vbArray( dict.Item( "length") - 1)
	Dim I
	For I = 0 To dict.Item( "length") - 1
		vbArray( I) = dict.Item( I)
	Next
	Dict2VBArray = vbArray
End Function
