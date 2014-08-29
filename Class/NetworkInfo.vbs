Option Explicit
Class NetworkInfo

	Private networkDict_

	Public Sub Class_Initialize()
	
		Call setNetworkStatus()
	
	End Sub

	Private Sub setNetworkStatus()

		
		Dim oLocator
		Dim oService
		Dim networkAdapts
		Dim networkAdapters
		Dim networkConfigs
		Dim adapterDict
		Dim configDict

		Set oLocator        = WScript.CreateObject("WbemScripting.SWbemLocator")
		Set oService        = oLocator.ConnectServer
		Set networkAdapters = oService.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionID IS NOT NULL")
		Set networkConfigs  = oService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration")
		
		Set adapterDict = createAdapterDict(networkAdapters)
		Set configDict  = createConfigDict(networkConfigs)
		
		WScript.echo adapterDict.count
		WScript.echo configDict.count
		
		
		Call createNetworkDict(adapterDict,configDict)
		
	End Sub
	
	Public Sub display()
	
		WScript.echo "********************************************************************"
		WScript.echo "* Networks                                                         *"
		WScript.echo "********************************************************************"		
		WScript.echo ""
		
		Dim adapterInfo
		Dim count
		adapterDict_.Keys

		For Each networkId In networkDict 
			
			Dim conId
			Dim adapterName
			Dim mac
			Dim ip
			Dim subnet
			Dim dhcp
			
			conId       = networkDict_.Item(networkId).Item("conId")
			adapterName = networkDict_.Item(networkId).Item("adapterName")
			mac         = networkDict_.Item(networkId).Item("mac")
			ip          = networkDict_.Item(networkId).Item("ip")
			subnet	    = networkDict_.Item(networkId).Item("subnet")
			dhcp        = networkDict_.Item(networkId).Item("dhcp")
			
			WScript.echo "接続名　　 ： " & conId
			WScript.echo "アダプタ名 ： " & adapterName
			WScript.echo "Macアドレス： " & mac
			WScript.echo "IPアドレス ： " & ip
			WScript.echo "SubnetMask ： " & subnet
			WScript.echo "DHCP　　　 ： " & dhcp
			WScript.echo ""
		
		Next
	End Sub
	
	Private Function createAdapterDict(networkAdapters)

		Dim adapterDict
		Dim oneOfAdapter
		Dim count 
		
		Set adapterDict = CreateObject("Scripting.Dictionary")
		
		count = 0
		For Each oneOfAdapter In networkAdapters

			Dim tmpDict
			Dim mac
			Set tmpDict = CreateObject("Scripting.Dictionary")
			
			tmpDict.Add "conId"       , oneOfAdapter.NetConnectionID
			tmpDict.Add "adapterName" , oneOfAdapter.Name
			
			mac = oneOfAdapter.MACAddress
			If IsNull(mac) then
				mac = ""
			End If
			
			tmpDict.Add "mac" , mac
			
			Dim adapterId 
			adapterId = "adapter" & count
			adapterDict.Add adapterId , tmpDict
			
			count = count + 1
		
		Next
		
		Set createAdapterDict = adapterDict
		
	End Function
	
	Private Function createConfigDict(networkConfigs)
		
		Dim configDict
		Dim oneOfConfig
		Dim count
		Set configDict = CreateObject("Scripting.Dictionary")
		
		count = 0
		For Each oneOfConfig In networkConfigs
		
			Dim tmpDict
			Set tmpDict = CreateObject("Scripting.Dictionary")
			
			tmpDict.Add "adapterName" , oneOfConfig.Description
			tmpDict.Add "dhcp"        , CStr(oneOfConfig.DHCPEnabled)
			
			If oneOfConfig.IPEnabled = True Then 
				tmpDict.Add "ip"     , oneOfConfig.IPAddress(0)
				tmpDict.Add "subnet" , oneOfConfig.IPSubnet(0)
			Else
				tmpDict.Add "ip"     , ""
				tmpDict.Add "subnet" , ""
			End If
			
			Dim configId 
			configId = "config" & count
			configDict.Add configId , tmpDict
			
			count = count + 1
		
		Next
		
		Set createConfigDict = configDict
		
	End Function
	
	
	Private Sub createNetworkDict(adapterDict,configDict)
		
		Dim count
		Dim oneOfAdapter
		count = 0
		For Each oneOfAdapter In adapterDict
			
			Dim tmpDict
			Dim a_adapterName
			Set tmpDict       = CreateObject("Scripting.Dictionary")
			a_adapterName     = adapterDict.Item("oneOfAdapter").Item("adapterName")
			
			For Each oneOfConfig In configDict
			
				Dim c_adapterName
				c_adapterName = adapterDict.Item("oneOfConfig").Item("adapterName")
				If StrComp(a_adapterName , c_adapterName) Then
				
					tmpDict.Add "conId"       , adapterDict.Item("oneOfAdapter").Item("conId")
					tmpDict.Add "adapterName" , a_adapterName
					tmpDict.Add "mac"         , adapterDict.Item("oneOfAdapter").Item("mac")
					tmpDict.Add "ip"          , adapterDict.Item("oneOfConfig").Item("ip")
					tmpDict.Add "subnet"      , adapterDict.Item("oneOfConfig").Item("subnet")
					tmpDict.Add "dhcp"        , adapterDict.Item("oneOfConfig").Item("dhcp")
				
				End If
			
			Next
			
			Dim networkId
			networkId = "network" & count
			networkDict_.Add networkId , tmpDict  
			count = count + 1
			
		Next
	End Sub
	
End Class


Dim test
Set test = New NetworkInfo
