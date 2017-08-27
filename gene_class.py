adDict = {
    "String": "adVarChar",
    "Boolean": "AdBoolean",
    "Integer": "adInteger",
    "Double": "adDouble",
    "Single": "adSingle",
}


def gene_class(properties, class_name, file_path):
    if not isinstance(properties, list):
        return
    def_str = "Public Hash As String\n"
    for prop in properties:
        if not isinstance(prop, list) or len(prop) < 3 or prop[1] not in adDict:
            return
        def_str += "Public %s As %s\n" % (prop[0], prop[1])
    
    decode_str = """
Public Sub DecodeRs(ByRef recordset As ADODB.recordset)
    Hash = recordset("Hash").Value
    """
    for prop in properties:
        decode_str += "%s = recordset(\"%s\").Value\n" % (prop[0], prop[2])
    decode_str += "End Sub\n"

    seek_str = """
Public Function SeekByHash(ByRef conn As ADODB.Connection, ByRef hashStr As String) As ADODB.recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select * from %s where Hash=@Hash"
    cmd.Parameters.Append cmd.CreateParameter("@Hash", adVarChar, adParamInput, Len(hashStr), hashStr)
    Set SeekByHash = cmd.Execute
End Function
""" % class_name

    prop_strs = []
    for prop in properties:
        if prop[1] == "String":
            prop_strs.append(prop[0])
        else:
            prop_strs.append("Str(%s)" % prop[0])
    hash_str = """
Private Sub GetHash()
    Hash = SHA1HASH(%s)
End Sub
""" % " & ".join(prop_strs)

    insert_str = """
Private Sub InsertDef(ByRef conn As ADODB.Connection)
    If Hash = "" Then
        GetHash
    End If
    
    Dim keys, values
    keys = Array("Hash\""""

    for prop in properties:
        insert_str += ", \"%s\"" % prop[2]
    
    insert_str += """)
    values = Array("@p1\""""

    for i in range(2, 2 + len(properties)):
        insert_str += ", \"@p%d\"" % i
    
    insert_str += """)
    Dim keyString, valueString, commandString As String
    commandString = "INSERT INTO %s (" & Join(keys, ",") & _
    ") VALUES (" & Join(values, ",") & ")"
    
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = commandString

    cmd.Parameters.Append cmd.CreateParameter("@p1", adVarChar, adParamInput, Len(Hash), Hash)
    """ % class_name
    
    for i, prop in enumerate(properties):
        if prop[1] == "String":
            insert_str += "cmd.Parameters.Append cmd.CreateParameter(\"@p%d\", %s, adParamInput, Len(%s), %s)\n    " % (i+2, adDict[prop[1]], prop[0], prop[0])
        else:
            insert_str += "cmd.Parameters.Append cmd.CreateParameter(\"@p%d\", %s, adParamInput, , %s)\n    " % (i+2, adDict[prop[1]], prop[0])
    insert_str += """cmd.Execute
End Sub
"""

    seek_update_str = """
Public Sub SeekAndUpdate(ByRef conn As ADODB.Connection)
    If Hash = "" Then
        GetHash
    End If
    
    Dim rs As ADODB.recordset
    Set rs = SeekByHash(conn, Hash)
    If rs.EOF Then
        InsertDef conn
    End If
    rs.Close
End Sub"""

    with open(file_path, "w") as fid:
        fid.write(def_str + decode_str + seek_str + hash_str + insert_str + seek_update_str)


if __name__ == "__main__":
    props = [
        ["Command_Name", "String", "Command_Name"],
        ["Factor", "String", "Factor"],
        ["Byte1", "String", "Byte1"],
        ["Bytes", "String", "Bytes"],
        ["Bytes3", "String", "Bytes3"],
    ]
    gene_class(props, "Rule_CMD", "~tmp.txt")