' ########################################################################################
' Microsoft Windows
' File: AfxDPIProcs2.inc
' Contents: Windows High DPI wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"

NAMESPACE Afx2


' ========================================================================================
' Macro for debug
' To allow debugging, define _DWSTRING_LIST_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _DWSTRING_LIST_DEBUG_
   #define _DWSTRING_LIST_DEBUG_ 0
#ENDIF
#ifndef _DWSTRING_LIST_DP_
   #define _DWSTRING_LIST_DP_ 1
   #MACRO DWSTRING_LIST_DP(st)
      #IF (_DWSTRING_LIST_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' Implementation of a double linked-list for the DWSTRING data type
' Usage example:
' Function to create and return a pointer to the List
' ----------------------------------------------------------------------------------------
' FUNCTION ProcessData () AS DWSTRING_LIST PTR
'    DIM myList AS DWSTRING_LIST PTR = NEW DWSTRING_LIST
'    ' Simulating result generation
'    myList->AddNode(DWSTRING("Result 1"))
'    myList->AddNode(DWSTRING("Result 2"))
'    myList->AddNode(DWSTRING("Result 3"))
'    ' Return the pointer to the list
'    RETURN myList
' END FUNCTION
' ----------------------------------------------------------------------------------------
' Get the results
' DIM myListPtr AS DWSTRING_LIST PTR
' myListPtr = ProcessData()
' ' Retrieve and print the results
' IF myListPtr <> NULL THEN
'    ' Iterate forward
'    OutputDebugStringW ("Iterating forward:")
'    DIM currentNode AS DWSTRING_NODE PTR = myListPtr->head
'    WHILE currentNode <> NULL
'       PRINT currentNode->dws
'       currentNode = currentNode->pNext
'    WEND
'    ' Iterate backward
'    OutputDebugStringW ("Iterating backward:")
'    currentNode = myListPtr->tail
'    WHILE currentNode <> NULL
'       PRINT currentNode->dws
'       currentNode = currentNode->pPrev
'    WEND
'    ' Delete the list, triggering cleanup automatically
'    Delete myListPtr
' END IF
' ========================================================================================

TYPE DWSTRING_NODE
   dws AS DWSTRING
   pNext AS DWSTRING_NODE PTR
   pPrev AS DWSTRING_NODE PTR
   ' Destructor to clean up memory
   DECLARE DESTRUCTOR
END TYPE

PRIVATE DESTRUCTOR DWSTRING_NODE
   ' DWSTRING's destructor will automatically handle cleanup
   DWSTRING_LIST_DP ("")
END DESTRUCTOR

TYPE DWSTRING_LIST
   head As DWSTRING_NODE PTR
   tail As DWSTRING_NODE PTR
   DECLARE SUB AddNode (BYREF newString AS DWSTRING)
   ' Destructor: Automatically cleans up when deleted
   DECLARE DESTRUCTOR
END TYPE

PRIVATE DESTRUCTOR DWSTRING_LIST
   DWSTRING_LIST_DP ("")
   DIM current AS DWSTRING_NODE PTR = head
   WHILE current <> NULL
      DIM nextNode AS DWSTRING_NODE PTR = current->pNext
      Delete current
      current = nextNode
   WEND
END DESTRUCTOR

   ' Add a new DWSTRING to the list
PRIVATE SUB DWSTRING_LIST.AddNode (BYREF newString AS DWSTRING)
   DWSTRING_LIST_DP ("")
   DIM newNode AS DWSTRING_NODE PTR = NEW DWSTRING_NODE
   newNode->dws = newString
   newNode->pNext = NULL
   newNode->pPrev = tail ' Link new node to the current tail
   IF tail <> NULL THEN
      tail->pNext = newNode
   ELSE
      head = newNode
   END IF
   tail = newNode
END SUB
' ========================================================================================

' ========================================================================================
' Implementation of an indexed double linked-list for the DWSTRING data type
' Usage example:
' 'Function to Create and Return a Pointer to the List
' ----------------------------------------------------------------------------------------
' FUNCTION ProcessData () AS DWSTRING_INDEXED_LIST PTR
'    DIM myList AS DWSTRING_INDEXED_LIST PTR = NEW DWSTRING_INDEXED_LIST
'    ' Simulating result generation
'    myList->AddNode(DWSTRING("Result 1"))
'    myList->AddNode(DWSTRING("Result 2"))
'    myList->AddNode(DWSTRING("Result 3"))
'    ' Return the pointer to the list
'    RETURN myList
' END FUNCTION
' ----------------------------------------------------------------------------------------
' ' Get the results
' DIM myListPtr AS DWSTRING_INDEXED_LIST PTR
' myListPtr = ProcessData()
' ' Retrieve and print the results
' IF myListPtr <> NULL THEN
'    ' Get the number of nodes
'    DIM count AS LONG = myListPtr->Count
'    ' Direct access using 0-based index
'    DIM node AS DWSTRING_INDEXED_NODE PTR
'    node = myListPtr->GetByIndex(1) ' Get second element
'    IF node <> NULL THEN PRINT node->dws
'    ' Delete the list, triggering cleanup automatically
'    Delete myListPtr
' END IF
' ========================================================================================

TYPE DWSTRING_INDEXED_NODE
   dws AS DWSTRING
   pNext AS DWSTRING_INDEXED_NODE PTR
   pPrev AS DWSTRING_INDEXED_NODE PTR
   ' Destructor to clean up memory
   DECLARE DESTRUCTOR
END TYPE

PRIVATE DESTRUCTOR DWSTRING_INDEXED_NODE
   ' DWSTRING's destructor will automatically handle cleanup
   DWSTRING_LIST_DP ("")
END DESTRUCTOR

TYPE DWSTRING_INDEXED_LIST
   head AS DWSTRING_INDEXED_NODE PTR
   tail AS DWSTRING_INDEXED_NODE PTR
   DIM index(ANY) AS DWSTRING_INDEXED_NODE PTR ' Array for indexed access
   count AS LONG ' Number of nodes
   DECLARE SUB AddNode (BYREF newString AS DWSTRING)
   DECLARE FUNCTION GetByIndex (BYVAL i AS LONG) AS DWSTRING_INDEXED_NODE PTR
   DECLARE FUNCTION GetCount () AS LONG
   ' Destructor: Automatically cleans up when deleted
   DECLARE DESTRUCTOR
END TYPE

PRIVATE DESTRUCTOR DWSTRING_INDEXED_LIST
   DWSTRING_LIST_DP ("")
   DIM current AS DWSTRING_INDEXED_NODE PTR = head
   WHILE current <> NULL
      DIM nextNode AS DWSTRING_INDEXED_NODE PTR = current->pNext
      Delete current
      current = nextNode
   WEND
   ERASE index ' Cleanup index array
END DESTRUCTOR

   ' Add a new DWSTRING to the list
PRIVATE SUB DWSTRING_INDEXED_LIST.AddNode (BYREF newString AS DWSTRING)
   DWSTRING_LIST_DP ("")
   DIM newNode AS DWSTRING_INDEXED_NODE PTR = NEW DWSTRING_INDEXED_NODE
   newNode->dws = newString
   newNode->pNext = NULL
   newNode->pPrev = tail ' Link new node to the current tail
   IF tail <> NULL THEN
      tail->pNext = newNode
   ELSE
      head = newNode
   END IF
   tail = newNode
   ' Update index
   count += 1
   REDIM PRESERVE index(count - 1)
   index(count - 1) = newNode
END SUB

' Get the number of nodes
PRIVATE FUNCTION DWSTRING_INDEXED_LIST.GetCount () AS LONG
   RETURN count
END FUNCTION

' Get node by index
PRIVATE FUNCTION DWSTRING_INDEXED_LIST.GetByIndex (BYVAL i AS LONG) AS DWSTRING_INDEXED_NODE PTR
   IF i >= 0 AND i < count THEN
      RETURN index(i)
   ELSE
      RETURN NULL ' Out of bounds
   END IF
END FUNCTION
' ========================================================================================

END NAMESPACE

