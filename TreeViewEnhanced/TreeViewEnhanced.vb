﻿
Option Strict On

Imports System.Windows.Forms

Imports System.Drawing

Imports System.ComponentModel

Public Class TreeViewEnhanced
    Inherits TreeView
#Region "Events"
    Public Event NodeMovedByDrag As EventHandler(Of NodeMovedByDragEventArgs)
    Protected Overridable Sub OnNodeMovedByDrag(ByVal e As NodeMovedByDragEventArgs)
        RaiseEvent NodeMovedByDrag(Me, e)
    End Sub

    Public Event NodeMovingByDrag As EventHandler(Of NodeMovingByDragEventArgs)
    Protected Overridable Sub OnNodeMovingByDrag(ByVal e As NodeMovingByDragEventArgs)
        RaiseEvent NodeMovingByDrag(Me, e)
    End Sub

    Public Event NodeDraggingOver As EventHandler(Of NodeDraggingOverEventArgs)
    Protected Overridable Sub OnNodeDraggingOver(ByVal e As NodeDraggingOverEventArgs)
        RaiseEvent NodeDraggingOver(Me, e)
    End Sub

    'Public Event RemoveChecked As EventHandler
    'Protected Overridable Sub OnRemoveChecked(ByVal e As EventArgs)
    '    RaiseEvent RemoveChecked(Me, e)
    'End Sub

#End Region

    Private _NextTransition As String

    Sub New()
        MyBase.AllowDrop = False
        MyBase.BackColor = Color.Wheat
        _NextTransition = "Expand"
        AddHandler MyBase.AfterCheck, AddressOf Me.AfterCheck
 
        Dim contextmenu As New ContextMenu()
 
        Dim expandTree As New MenuItem("Expand")
        AddHandler expandTree.Click, AddressOf expandIt
        contextmenu.MenuItems.Add(expandTree)

        Dim contractTree As New MenuItem("Contract")
        AddHandler contractTree.Click, AddressOf contractIt
        contextmenu.MenuItems.Add(contractTree)

        Dim collapseTree As New MenuItem("Collapse")
        AddHandler collapseTree.Click, AddressOf collapseIt
        contextmenu.MenuItems.Add(collapseTree)

        contextmenu.MenuItems.Add("-")

        Dim removeTicked As New MenuItem("Remove Ticked")
        AddHandler removeTicked.Click, AddressOf RemoveChecked
        contextmenu.MenuItems.Add(removeTicked)

        Dim removeUnticked As New MenuItem("Remove Unticked")
        AddHandler removeUnticked.Click, AddressOf RemoveUnchecked
        contextmenu.MenuItems.Add(removeUnticked)
 
        MyBase.ContextMenu = (contextmenu)
 
    End Sub

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Shadows ReadOnly Property DrawMode() As System.Windows.Forms.TreeViewDrawMode
        Get
            Return MyBase.DrawMode
        End Get
    End Property

    'Decided to expose the AllowDrop Property, so we can use 
    ''<DefaultValue(True)> _
    '<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    'Public Shadows ReadOnly Property AllowDrop() As Boolean
    '    Get
    '        Return MyBase.AllowDrop
    '    End Get
    'End Property

    '' expose the AllowDrop Property 
    'Public Overrides Property AllowDrop As Boolean
    '    Get
    '        Return MyBase.AllowDrop
    '    End Get
    '    Set(value As Boolean)
    '        MyBase.AllowDrop = value
    '    End Set
    'End Property
 

    Protected Overrides Sub OnItemDrag(ByVal e As System.Windows.Forms.ItemDragEventArgs)
        MyBase.DoDragDrop(e.Item, DragDropEffects.Move)

        MyBase.OnItemDrag(e)
    End Sub

    Protected Overrides Sub OnDragOver(ByVal drgevent As System.Windows.Forms.DragEventArgs)

        'Get the node we are currently over
        'while dragging another node
        Dim targetNode As TreeNode = MyBase.GetNodeAt(MyBase.PointToClient(New Point(drgevent.X, drgevent.Y)))

        'Get the node being dragged
        Dim dragNode As TreeNode = FindNodeInDataObject(drgevent.Data)

        Dim eaDraggingOver As New NodeDraggingOverEventArgs(dragNode, targetNode)
        OnNodeDraggingOver(eaDraggingOver)

        If eaDraggingOver.DropIsLegal = False Then
            drgevent.Effect = DragDropEffects.None
            Return
        End If


        'If we are not currently dragging over
        'a node...
        If targetNode Is Nothing Then

            'Let no node be selected
            MyBase.SelectedNode = Nothing

            'Allow the move because its valid
            'to drag a node over the TreeView itself
            'the drop will place the node being dragged
            'in the root
            drgevent.Effect = DragDropEffects.Move

            'Get out
            Return
        End If

        'This would only be nothing if something is being
        'dragged over the TreeView that isn't a node
        If dragNode IsNot Nothing Then

            'Illegal to drop nodes inside their descendants
            'Its not logical
            If targetNode Is dragNode OrElse IsNodeDescendant(targetNode, dragNode) Then

                'Prevents a drop
                drgevent.Effect = DragDropEffects.None
            Else

                'Allows a drop
                drgevent.Effect = DragDropEffects.Move
                MyBase.SelectedNode = targetNode
            End If

        End If

        MyBase.OnDragOver(drgevent)
    End Sub

    Protected Overrides Sub OnDragDrop(ByVal drgevent As System.Windows.Forms.DragEventArgs)
        Dim dragNode As TreeNode = FindNodeInDataObject(drgevent.Data)

        If dragNode IsNot Nothing Then
            'Dim dragNode As TreeNode = DirectCast(drgevent.Data.GetData(GetType(TreeNode)), TreeNode)

            'Get the parent of the node before moving
            Dim prevParent As TreeNode = dragNode.Parent
            Dim parentToBe As TreeNode = If(MyBase.SelectedNode Is Nothing, Nothing, MyBase.SelectedNode)

            Dim eaNodeMoving As New NodeMovingByDragEventArgs(dragNode, prevParent, parentToBe)

            OnNodeMovingByDrag(eaNodeMoving)

            If eaNodeMoving.CancelMove = False Then

                dragNode.Remove()
                If MyBase.SelectedNode IsNot Nothing Then
                    MyBase.SelectedNode.Nodes.Add(dragNode)
                Else
                    MyBase.Nodes.Add(dragNode)
                End If
                OnNodeMovedByDrag(New NodeMovedByDragEventArgs(dragNode, prevParent))
            End If

        End If


        MyBase.OnDragDrop(drgevent)
    End Sub

    Private Function IsNodeDescendant(ByVal node As TreeNode, ByVal potentialElder As TreeNode) As Boolean
        Dim n As TreeNode

        If node Is Nothing OrElse potentialElder Is Nothing Then Return False

        Do
            n = node.Parent

            If n IsNot Nothing Then
                If n Is potentialElder Then
                    Return True
                Else
                    node = n
                End If
            End If
        Loop Until n Is Nothing

        Return False
    End Function

    Private Function FindNodeInDataObject(ByVal dataObject As IDataObject) As TreeNode

        For Each format As String In dataObject.GetFormats
            Dim data As Object = dataObject.GetData(format)

            If GetType(TreeNode).IsAssignableFrom(data.GetType) Then
                Return DirectCast(data, TreeNode)
            End If
        Next

        Return Nothing
    End Function


    Private Sub CheckChildNodes(treeNode As TreeNode, nodeChecked As Boolean)
        Dim node As TreeNode
        For Each node In treeNode.Nodes
            node.Checked = nodeChecked
            If nodeChecked Then
                node.ExpandAll()
            End If
        Next node
    End Sub

    Private Shadows Sub AfterCheck(sender As Object, e As TreeViewEventArgs)

        CheckChildNodes(e.Node, e.Node.Checked)

    End Sub


    Public Shared Sub TickNode(ByRef givenNodes As TreeNodeCollection,
                               ByVal search As String,
                               ByRef found As Boolean,
                               Optional ByVal matchPath As Boolean = False,
                               Optional ByVal path As String = "",
                               Optional ByVal checked As Boolean = True)
        'Console.WriteLine("TickNode:" & search & ":" & path)
        Dim node As TreeNode
        For Each node In givenNodes
            'Console.WriteLine(node.Text)
            If Not found Then
                If (matchPath And path & node.Text = search) Or
                   (Not matchPath And (node.Text = search Or node.Name = search Or node.Tag Is search)) Then 'WHY DOES THIS USES "Is"? The Is operator determines if two object references refer to the same object. SUSPECT TYPO

                    node.Checked = checked
                    found = True

                End If

                TickNode(node.Nodes, search, found, matchPath, path & node.Text & "\", checked)
            End If
        Next

    End Sub

    Public Sub TickNode(ByVal search As String,
                        ByRef found As Boolean,
                        Optional ByVal matchPath As Boolean = False,
                        Optional ByVal path As String = "",
                        Optional ByVal checked As Boolean = True)

        TickNode(givenNodes:=MyBase.Nodes,
                 search:=search,
                 found:=found,
                 matchPath:=matchPath,
                 path:=path,
                 checked:=checked)

    End Sub

    Public Sub TickNodes(ByVal tickList As Collection,
                         Optional ByVal matchPath As Boolean = False,
                         Optional ByVal checked As Boolean = True)

        'Tick Nodes in the treeview
        For Each tickItem In tickList
            Dim found As Boolean = False
            TickNode(search:=tickItem.ToString,
                     found:=found,
                     matchPath:=matchPath,
                     path:="",
                     checked:=checked)

        Next

    End Sub

    Public Sub TickNodesByValue(ByVal tickList As Dictionary(Of String, String),
                                Optional ByVal matchPath As Boolean = True,
                                Optional ByVal checked As Boolean = True)

        'Match on the path

        'Tick Nodes in the treeview
        For Each tickItem In tickList
            Dim found As Boolean = False
            TickNode(search:=tickItem.Value,
                     found:=found,
                     matchPath:=matchPath,
                     path:="",
                     checked:=checked)

        Next

    End Sub

    '@TODO THIS PROBABLY WILL NEED TO CHANGE TO MATCH ON NODE.NAME ONLY.
    Public Sub TickNodesByKey(ByVal tickList As Dictionary(Of String, String),
                              Optional ByVal matchPath As Boolean = False,
                              Optional ByVal checked As Boolean = True)

        'Match on the name (key) of the leaf node value only, not the full path.

        'Tick Nodes in the treeview
        For Each tickItem In tickList
            Dim found As Boolean = False
            TickNode(search:=tickItem.Key,
                     found:=found,
                     matchPath:=matchPath,
                     path:="",
                     checked:=checked)

        Next

    End Sub

    'Not Used in GitPatcher
    Public Shared Sub RenameNode(ByRef givenNodes As TreeNodeCollection,
                                 ByVal search As String,
                                 ByVal newName As String,
                                 ByRef found As Boolean,
                                 Optional ByVal matchPath As Boolean = False,
                                 Optional ByVal path As String = "",
                                 Optional ByVal checked As Boolean = True)
        'Console.WriteLine("RenameNode:" & search & ":" & path)
        Dim node As TreeNode
        For Each node In givenNodes
            'Console.WriteLine(node.Text)
            If Not found Then
                If (matchPath And path & node.Text = search) Or
                   (Not matchPath And (node.Text = search Or node.Tag.ToString = search)) Then

                    node.Name = newName
                    found = True

                End If

                RenameNode(givenNodes:=node.Nodes,
                           search:=search,
                           newName:=newName,
                           found:=found,
                           matchPath:=matchPath,
                           path:=path & node.Text & "\",
                           checked:=checked)
            End If
        Next

    End Sub

    'Not Used in GitPatcher
    Public Sub RenameNode(ByVal search As String,
                                 ByVal newName As String,
                                 ByRef found As Boolean,
                                 Optional ByVal matchPath As Boolean = False,
                                 Optional ByVal path As String = "",
                                 Optional ByVal checked As Boolean = True)

        RenameNode(givenNodes:=MyBase.Nodes,
                   search:=search,
                   newName:=newName,
                   found:=found,
                   matchPath:=matchPath,
                   path:=path,
                   checked:=checked)

    End Sub


    Public Shared Sub RemoveNodes(ByRef givenNodes As TreeNodeCollection, ByVal checked As Boolean)
        Dim node As TreeNode
        For i As Integer = givenNodes.Count - 1 To 0 Step -1

            node = givenNodes(i)

            If node.Nodes.Count > 0 Then
                'non-leaf
                RemoveNodes(node.Nodes, checked)
                'if no leaves left, remove the branch
                If node.Nodes.Count = 0 Then
                    givenNodes.Remove(node)
                End If

            Else
                'leaf
                If node.Checked = checked Then
                    givenNodes.Remove(node)
                End If
            End If

        Next

    End Sub


    Public Sub RemoveNodes(ByVal checked As Boolean)

        RemoveNodes(MyBase.Nodes, checked)

    End Sub

    Public Sub RemoveItem(ByRef givenNodes As TreeNodeCollection,
                          ByVal item As String,
                  Optional ByVal matchText As Boolean = False,
                  Optional ByVal matchTag As Boolean = False)



        For i As Integer = givenNodes.Count - 1 To 0 Step -1

            Dim node As TreeNode = givenNodes(i)

            If node.Nodes.Count > 0 Then
                'non-leaf
                RemoveItem(node.Nodes, item, matchText, matchTag)
                'if no leaves left, remove the branch
                If node.Nodes.Count = 0 Then
                    givenNodes.Remove(node)
                End If

            Else
                'leaf
                If (matchText AndAlso node.Text = item) Or
                   (matchTag AndAlso node.Tag.ToString = item) Then
                    givenNodes.RemoveAt(i)
                    'removedItem = True
                    'Exit For 'assume there is a unique list, so finish searching once we have it.
                End If
            End If

        Next i

    End Sub

    Public Sub RemoveItem(ByVal item As String,
                  Optional ByVal matchText As Boolean = False,
                  Optional ByVal matchTag As Boolean = False)

        RemoveItem(MyBase.Nodes, item, matchText, matchTag)

    End Sub

    Public Sub RemoveItems(ByRef items As Collection,
                  Optional ByVal matchText As Boolean = False,
                  Optional ByVal matchTag As Boolean = False)

        For Each item As String In items
            RemoveItem(item, matchText, matchTag)
        Next item

    End Sub


    Sub collapseIt(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Me.CollapseAll()
    End Sub

    Sub expandIt(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Me.ExpandAll()
    End Sub

    Sub contractIt(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Me.showCheckedNodes()
    End Sub

    Sub RemoveChecked(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Me.RemoveNodes(True)
    End Sub

    Sub RemoveUnchecked(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Me.RemoveNodes(False)
    End Sub

    Private Shared Sub ReadLeafNodes(ByRef givenNodes As TreeNodeCollection, ByRef fullPathsList As Collection, Optional ByVal checkedOnly As Boolean = False)
        'Recursive tree search
        'Param checkedOnly - when true will ReadCheckedLeafNodes
        Dim node As TreeNode
        For Each node In givenNodes

            If node.Nodes.Count = 0 Then
                'Leaf node
                If (Not checkedOnly Or node.Checked) Then

                    If node.Name = "" Then
                        If Not fullPathsList.Contains(node.FullPath) Then
                            fullPathsList.Add(node.FullPath, node.FullPath)
                        End If
                        'Node 
                    Else
                        'Node name is populated so using this as the key
                        If Not fullPathsList.Contains(node.Name) Then
                            fullPathsList.Add(Item:=node.FullPath,
                                              Key:=node.Name)
                        End If

                    End If


                End If

            Else
                ReadLeafNodes(node.Nodes, fullPathsList, checkedOnly)

            End If

        Next

    End Sub


    Public Sub ReadCheckedLeafNodes(ByRef fullPathsList As Collection)

        ReadLeafNodes(MyBase.Nodes, fullPathsList, True)

    End Sub

    Public Sub ReadAllLeafNodes(ByRef fullPathsList As Collection)

        ReadLeafNodes(MyBase.Nodes, fullPathsList, False)

    End Sub


    Private Shared Sub ReadLeafNodes(ByRef givenNodes As TreeNodeCollection, ByRef fullPathsList As Dictionary(Of String, String), Optional ByVal checkedOnly As Boolean = False)
        'Recursive tree search
        'Param checkedOnly - when true will ReadCheckedLeafNodes
        Dim node As TreeNode
        For Each node In givenNodes

            If node.Nodes.Count = 0 Then
                'Leaf node
                If (Not checkedOnly Or node.Checked) Then

                    If node.Name = "" Then
                        If Not fullPathsList.ContainsKey(node.FullPath) Then
                            fullPathsList.Add(node.FullPath, node.FullPath)
                        End If
                        'Node 
                    Else
                        'Node name is populated so using this as the key
                        If Not fullPathsList.ContainsKey(node.Name) Then
                            fullPathsList.Add(key:=node.Name,
                                              value:=node.FullPath)
                        End If

                    End If


                End If

            Else
                ReadLeafNodes(node.Nodes, fullPathsList, checkedOnly)

            End If

        Next

    End Sub


    Public Sub ReadCheckedLeafNodes(ByRef fullPathsList As Dictionary(Of String, String))

        ReadLeafNodes(MyBase.Nodes, fullPathsList, True)

    End Sub

    Public Sub ReadAllLeafNodes(ByRef fullPathsList As Dictionary(Of String, String))

        ReadLeafNodes(MyBase.Nodes, fullPathsList, False)

    End Sub


    Private Shared Function getFirstSegment(ByVal ipath As String, ByVal idelim As String) As String

        Return ipath.Split(CChar(idelim))(0)
    End Function


    Private Shared Function dropFirstSegment(ByVal ipath As String, ByVal idelim As String) As String

        Dim l_from_first As String = Nothing
        Dim delim_pos As Integer = ipath.IndexOf(idelim)
        If delim_pos > 0 Then
            l_from_first = ipath.Remove(0, delim_pos + 1)
        End If

        Return l_from_first
    End Function


    Public Shared Function AddNode(ByRef nodes As TreeNodeCollection,
                                   ByVal fullPath As String,
                                   ByVal remainderPath As String,
                                   Optional ByVal delim As String = "\",
                                   Optional ByVal checked As Boolean = False,
                                   Optional ByVal data As String = "") As Boolean
        'Commented out the logging, as it severely reduces performance.
        Dim first_segment As String = getFirstSegment(remainderPath, delim)
        Dim remainder As String = dropFirstSegment(remainderPath, delim)

        Dim lFound As Boolean = False
        Dim node As TreeNode
        'First try to find the node
        For Each node In nodes
            If fullPath = node.FullPath Then
                'If node.FullPath.ToString = fullPath Then
                'Yay found it nothing to do
                Return True
                'Node Full path must match first part of given full path, and current node must match exactly current segment
            ElseIf InStr(fullPath, node.FullPath.ToString) = 1 And first_segment = node.Text Then
                'Found a parent node at least, lets look for children
                lFound = AddNode(nodes:=node.Nodes,
                                 fullPath:=fullPath,
                                 remainderPath:=remainder,
                                 delim:=delim,
                                 checked:=checked,
                                 data:=data)
            End If

        Next

        If Not lFound And Not String.IsNullOrEmpty(first_segment) Then
            'Need to make a node
            Dim newNode As TreeNode = New TreeNode(first_segment)
            newNode.Tag = first_segment 'PAB not sure, but i think Tag represents the value of just the corresponding component of the fullpath.
            newNode.Checked = checked
            'newNode.Name = data
            nodes.Add(newNode)
            'If newNode.FullPath = fullPath Then
            If String.IsNullOrEmpty(remainder) Then
                'We made a leaf node!
                lFound = True
                newNode.Name = data 'add the data to the leaf only. This is to be used a the key.
            Else
                'Now follow this child
                lFound = AddNode(nodes:=newNode.Nodes,
                                 fullPath:=fullPath,
                                 remainderPath:=remainder,
                                 delim:=delim,
                                 checked:=checked,
                                 data:=data)

            End If

        End If
        If Not lFound Then
            MsgBox("Oops TreeView Node not found. Bad coding?")
        End If
        Return lFound


    End Function


    Public Function AddNode(ByVal fullPath As String,
                            Optional ByVal delim As String = "\",
                            Optional ByVal checked As Boolean = False,
                            Optional ByVal data As String = "") As Boolean

        Return AddNode(MyBase.Nodes,
                       fullPath:=fullPath,
                       remainderPath:=fullPath,
                       delim:=delim,
                       checked:=checked,
                       data:=data)

    End Function


    Public Sub populateTreeFromCollection(ByRef patches As Collection, Optional ByVal checked As Boolean = False, Optional clearNodes As Boolean = True)

        MyBase.PathSeparator = "\"
        If clearNodes Then
            MyBase.Nodes.Clear()
        End If

        'copy each item from listbox
        Dim found As Boolean
        Dim patch As String
        For Each patch In patches

            'find or create each node for item
            found = AddNode(fullPath:=patch,
                            delim:=MyBase.PathSeparator,
                            checked:=checked)

        Next

    End Sub

    Public Sub appendTreeFromCollection(ByRef patches As Collection, Optional ByVal checked As Boolean = False)

        populateTreeFromCollection(patches:=patches,
                                   checked:=checked,
                                   clearNodes:=False)

    End Sub

    Public Sub populateTreeFromDict(ByRef patches As Dictionary(Of String, String), 'object
                                    Optional ByVal checked As Boolean = False,
                                    Optional clearNodes As Boolean = True,
                                    Optional sortByKey As Boolean = False)

        MyBase.PathSeparator = "\"
        If clearNodes Then
            MyBase.Nodes.Clear()
        End If


        Dim found As Boolean


        ' Get list of keys.
        Dim keys As List(Of String) = patches.Keys.ToList

        If sortByKey Then
            ' Sort the keys.
            keys.Sort()
        End If


        Dim orderedCollection As Collection = New Collection()

        For Each key In keys

            'find or create each node for item
            found = AddNode(fullPath:=patches.Item(key).ToString,
                            delim:=MyBase.PathSeparator,
                            checked:=checked,
                            data:=key)
        Next

    End Sub



    Public Sub appendTreeFromDict(ByRef patches As Dictionary(Of String, String), 'object
                                  Optional ByVal checked As Boolean = False)

        populateTreeFromDict(patches:=patches,
                                   checked:=checked,
                                   clearNodes:=False)

    End Sub






    Public Sub showCheckedNodes()
        ' Disable redrawing of treeView1 to prevent flickering  
        ' while changes are made.
        MyBase.BeginUpdate()

        ' Collapse all nodes of treeView1.
        MyBase.CollapseAll()

        ' Add the CheckForCheckedChildren event handler to the BeforeExpand event. 
        AddHandler MyBase.BeforeExpand, AddressOf CheckForCheckedChildren

        ' Expand all nodes of treeView1. Nodes without checked children are  
        ' prevented from expanding by the checkForCheckedChildren event handler.
        MyBase.ExpandAll()

        ' Remove the checkForCheckedChildren event handler from the BeforeExpand  
        ' event so manual node expansion will work correctly. 
        RemoveHandler MyBase.BeforeExpand, AddressOf CheckForCheckedChildren

        ' Enable redrawing of treeView1.
        MyBase.EndUpdate()
    End Sub 'showCheckedNodesButton_Click

    ' Prevent expansion of a node that does not have any checked child nodes. 
    Shared Sub CheckForCheckedChildren(ByVal sender As Object, ByVal e As TreeViewCancelEventArgs)
        If Not HasCheckedChildNodes(e.Node) Then
            e.Cancel = True
        End If
    End Sub 'CheckForCheckedChildren

    ' Returns a value indicating whether the specified  
    ' TreeNode has checked child nodes. 
    Shared Function HasCheckedChildNodes(ByVal node As TreeNode) As Boolean
        If node.Nodes.Count = 0 Then
            Return False
        End If
        Dim childNode As TreeNode
        For Each childNode In node.Nodes
            If childNode.Checked Then
                Return True
            End If
            ' Recursively check the children of the current child node. 
            If HasCheckedChildNodes(childNode) Then
                Return True
            End If
        Next childNode
        Return False
    End Function 'HasCheckedChildNodes




End Class

Public Class NodeDraggingOverEventArgs
    Inherits EventArgs

    Private _DropLegal As Boolean
    Private _MovingNode As TreeNode
    Private _TargetNode As TreeNode

    Public Sub New(ByVal movingNode As TreeNode, ByVal targetNode As TreeNode)
        _DropLegal = True
        _MovingNode = movingNode
        _TargetNode = targetNode
    End Sub

    Public ReadOnly Property TargetNode() As TreeNode
        Get
            Return _TargetNode
        End Get
    End Property
    Public ReadOnly Property MovingNode() As TreeNode
        Get
            Return _MovingNode
        End Get
    End Property

    'Use this to disallow a drop
    Public Property DropIsLegal() As Boolean
        Get
            Return _DropLegal
        End Get
        Set(ByVal value As Boolean)
            _DropLegal = value
        End Set
    End Property


End Class

Public Class NodeMovingByDragEventArgs
    Inherits EventArgs

    Private _MovingNode As TreeNode
    Private _CurParent As TreeNode
    Private _ParentToBe As TreeNode

    Private _CancelMove As Boolean

    Public Sub New(ByVal nodeMoving As TreeNode, ByVal prevParent As TreeNode, ByVal parentToBe As TreeNode)
        _MovingNode = nodeMoving
        _CurParent = prevParent
        _ParentToBe = parentToBe
    End Sub
    Public Property CancelMove() As Boolean
        Get
            Return _cancelMove
        End Get
        Set(ByVal value As Boolean)
            _CancelMove = value
        End Set
    End Property
    Public ReadOnly Property MovingNode() As TreeNode
        Get
            Return _MovingNode
        End Get
    End Property
    Public ReadOnly Property CurrentParent() As TreeNode
        Get
            Return _CurParent
        End Get
    End Property
    Public ReadOnly Property ParentToBe() As TreeNode
        Get
            Return _ParentToBe
        End Get
    End Property
End Class

Public Class NodeMovedByDragEventArgs
    Inherits EventArgs

    Private _MovedNode As TreeNode
    Private _PreviousParent As TreeNode

    Public Sub New(ByVal nodeMoved As TreeNode, ByVal prevParent As TreeNode)
        _MovedNode = nodeMoved
        _PreviousParent = prevParent
    End Sub
    Public ReadOnly Property MovedNode() As TreeNode
        Get
            Return _MovedNode
        End Get
    End Property
    Public ReadOnly Property PreviousParent() As TreeNode
        Get
            Return _PreviousParent
        End Get
    End Property

End Class

