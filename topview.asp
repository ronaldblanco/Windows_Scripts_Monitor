
<!--
	Whatsup Gold
	
	Top View
	----------------
	main page. Displays a list of open
	maps and map status information.
	
-->


<HTML>

<%include% StandardPageHeader.asp>

<HEAD>
	<TITLE>WhatsUp Gold - Top View</TITLE>
	<META http-equiv="REFRESH" content="<%APPLICATION_SETTINGS% WEB_REFRESH_RATE>">
	
</HEAD>
<BODY>
<H1> <%APPLICATION_SETTINGS% MAIN_TITLE> </H1>

<%IF% IS_EVALUATION_VERSION>

	<%IF% (%GREATERTHAN% (%EVALUATION% DAYS_LEFT, 0))>
		<H3> Evaluation version - <%EVALUATION% DAYS_LEFT>&nbsp;day(s) remaining </H3>
	<%ENDIF%>

<%ENDIF%>

<%INCLUDE% MainPagePrefix.asp>

<TABLE cellSpacing=0 cellPadding=3 width="98%" bgColor="#c0c0c0" border=2>

	<!--MODIFICATION-->
	<TR><TD bgColor="#0080c0">User Process Monitor</TD></TR>
	<TR><TD bgcolor="#bbff77"><%include% Y:\00temp\SQL_STATE_share\PROCESSSTATUS.asp></TD></TR>
	<TR><TD bgcolor="#bbff77"><%include% Y:\00temp\SQL_STATE_share\PROCESSSTATUSCLIENT.asp></TD></TR>
	<TR><TD bgcolor="#ff0077">NEWMBA SQL:<%include% Y:\00temp\SQL_STATE_share\shareSQL_STATE\netstat.txt.more></TD></TR>
	<!--END MODIFICATION-->

	<TR>
	<td align=left bgColor="#0080c0" colSpan=2>
      <TABLE width="100%" border=0>
        <TR>
          <TD><%APPLICATION_SETTINGS% MAIN_TITLE></TD>
          <TD align=right>
		<%IF% (%MATCH% (%URL_VALUE% ("ViewMode",0),"Detail"))>
		  <A target=_NEW href="h_detailview.htm">Help</A>
		<%ELSE%>
		  <A target=_NEW href="h_top.htm">Help</A>
		<%ENDIF%>
		</TD></TR>
		</TABLE>
	</td>
	</TR>
		  
		  
  <TR>
<%IF% (%MATCH% (%URL_VALUE% ("ViewMode",0),"Detail"))>
    <TD vAlign=top width="98%">
      <TABLE cellSpacing=2 cellPadding=0 width="100%" bgColor="#c0c0c0" border=0>
        <TR>
          <TD width="20%"></TD>
          <TD width="40%"></TD>
          <TD width="40%"></TD>
 		</TR>
    <TR>
      <TD><hr size="4"></TD>
      <TD><hr size="4"></TD>
      <TD><hr size="4"></TD>
 	</TR>
	  <%START_LOADED_MAP_LIST%>
		<tr>
		<td><a href="map.asp?map=<%URL_ENCODE% (%MAP% FILENAME)>"> <%MAP% NAME> </A></td>
		<td>
		<table>
		<tr>
		<%IF% (%GREATERTHAN% (%MAP% TOTAL_DEVICES_UP,0))>
			<td bgcolor="#77ff77"><%MAP% TOTAL_DEVICES_UP>&nbsp;device(s) up</td>
		<%ELSE%>
			<td>No devices up!</td>
		<%ENDIF%>
		
		<%IF% (%GREATERTHAN% (%MAP% (TOTAL_DEVICES_DOWN),0))>
			<td bgcolor="#FF7777"><%MAP% TOTAL_DEVICES_DOWN>&nbsp;device(s) down</td>
		<%ELSE%>
			<td>No devices down</td>
		<%ENDIF%>
		</tr>
		</table>
		</td>
		
		<%IF% (%GREATERTHAN% (%MAP% (TOTAL_SERVICES_DOWN),0))>
			<td bgcolor="#ff77ff"><%MAP% TOTAL_SERVICES_DOWN>&nbsp;service(s) down</td>
		<%ELSE%>
			<td>No Services down</td>
		<%ENDIF%>
		</tr>

			<%START_DEVICE_LIST%>		  
			<TR>
			  <TD>&nbsp;</TD>
			  <TD><hr></TD>
			  <TD><hr></TD>
 			</TR>
			<TR>
			  <TD><!-- Leave Map Blank --></TD> 
 			  <TD>
			<%DEVICE% DISPLAY_NAME>

			<%IF% HAS_DEVICE_STATE_CHANGED>
				* 
			<%ENDIF%>

			<%IF% IS_MAP_ACCESS_HOST>
				&nbsp;(<A href="device.asp?map=<%URL_ENCODE%(%MAP% FILENAME)>&device=<%DEVICE% UNIQUE_ID>&MapViewMode=Summary"><%DEVICE% IP_ADDRESS></A>)
			<%ELSE%>
				&nbsp;(<%DEVICE% IP_ADDRESS>)
			<%ENDIF%>
			</td>


   			  <TD bgColor="<%DEVICE% POLL_STATE_COLOR>"><%DEVICE% POLL_TYPE></TD>
		</TR>

		<%START_SERVICE_STATISTICS_LIST%>
		<TR>
			  <TD><!-- Leave Map Blank --></TD> 
			  <TD><!-- Leave Name Blank --></TD> 
   			  <TD bgColor="<%SERVICE_STATISTICS% STATE_COLOR>"><%SERVICE_STATISTICS% FULLNAME></TD>
		</tr>
		<%END_SERVICE_STATISTICS_LIST%>
			  
		<%IF% (%GREATERTHAN% (%DEVICE% ALERT_COUNT, 0))>
			<TR>
			  <TD>&nbsp;</TD>
			  <TD>&nbsp;</TD>
			  <TD><hr></TD>
 			</TR>
			<%START_ALERT_LIST%>
				<TR>
					  <TD><!-- Leave Map Blank --></TD> 
					  <TD><!-- Leave Name Blank --></TD> 
					  <TD><%DEVICE_ALERT% ALERT_NAME></TD> 
				</TR>
			<%END_ALERT_LIST%>
		<%ENDIF%>
			  
	<%END_DEVICE_LIST%>		  
    <TR>
      <TD>&nbsp;</TD>
      <TD>&nbsp;</TD>
      <TD>&nbsp;</TD>
 	</TR>
    <TR>
      <TD><hr size="4"></TD>
      <TD><hr size="4"></TD>
      <TD><hr size="4"></TD>
 	</TR>

	
	<%END_LOADED_MAP_LIST%>
	 </TABLE>
	 </TD>
<%ELSE%>
    <TD vAlign=top width="98%">
      <TABLE cellSpacing=2 cellPadding=0 width="100%" bgColor="#c0c0c0" border=0>
        <TR>
          <TD width="40%"></TD>
          <TD width="20%"></TD>
          <TD width="20%"></TD>
          <TD width="20%"></TD>
		</TR>
        <TR>
          <TH align=left>Map</TH>
          <TH align=right>Items<BR>Up</TH>
          <TH align=right>Items<BR>Down</TH>
          <TH align=right>Items with<br>Services Down</TH></TR>
	  <%START_LOADED_MAP_LIST%>
		<tr >
		<td><a href="map.asp?map=<%URL_ENCODE% (%MAP% FILENAME)>"> <%MAP% NAME> </A></td>
		<%IF% (%GREATERTHAN% (%MAP% TOTAL_DEVICES_UP,0))>
			<td bgcolor="#77ff77" align=right><%MAP% TOTAL_DEVICES_UP></td>
		<%ELSE%>
			<td align=right><%MAP% TOTAL_DEVICES_UP></td>
		<%ENDIF%>
		
		<%IF% (%GREATERTHAN% (%MAP% (TOTAL_DEVICES_DOWN),0))>
			<td bgcolor="#FF7777" align=right><%MAP% TOTAL_DEVICES_DOWN></td>
		<%ELSE%>
			<td align=right> <%MAP% TOTAL_DEVICES_DOWN></td>
		<%ENDIF%>
		
		<%IF% (%GREATERTHAN% (%MAP% (TOTAL_SERVICES_DOWN),0))>
			<td bgcolor="#ff77ff" align=right><%MAP% TOTAL_SERVICES_DOWN></td>
		<%ELSE%>
			<td align=right><%MAP% TOTAL_SERVICES_DOWN></td>
		<%ENDIF%>
		</tr>
		
		<%END_LOADED_MAP_LIST%>
		  
	 </TABLE></TD>
		 
		 
    	<!-- Check if we need to play the sound file -->
		<!-- This will iterate over every map if necessary -->
		<%IF% HAS_ANY_MAP_STATE_CHANGED>
			<%IF% IS_SOUND_ENABLED>
			  <bgsound src="<%SOUND_TYPE% WEBDOWN>">
			<%ENDIF%>
		<%ENDIF%>

<%ENDIF%>

<%IF% (%MATCH% (%URL_VALUE% ("ViewMode",0),"Detail"))>
    <TD vAlign=top noWrap align=left>
	  <TABLE width="100%" border=1>
    	<TR>
          <TD noWrap align=center bgColor="#c0c0c0">
		  <A href="TopView.asp?ViewMode=Summary">Top View</A></TD></TR></TABLE>
<%ELSE%>
    <TD vAlign=top noWrap align=left>
	  <TABLE width="100%" border=1>
    	<TR>
          <TD noWrap align=center bgColor="#c0c0c0">
		  <A href="TopView.asp?ViewMode=Detail">Detail View</A></TD></TR></TABLE>
<%ENDIF%>

<%IF% IS_WEBSERVER_CONFIGURATION_ENABLED> <!-- APPSETTINGS CONFIGURE --> 
	<%IF% IS_USER_CONFIGURE_PROGRAM><!-- USER PROGRAM CONFIGURE -->
 	     <TABLE width="100%" border=1>
    	    <TR>
        	  <TD noWrap align=center bgColor="#c0c0c0">
   	  			<a href="AppSettingsForm.asp">Settings </A></TD></TR></TABLE>
	<!--IF% ACCESS_DEFAULT_MAP --> 
    		  <TABLE width="100%" border=1>
				<TR>
    			  <TD noWrap align=center bgColor="#c0c0c0">
				  <a href="NewMap.asp">New Map</A></TD></TR></TABLE>
	<!--ENDIF%-->
	      <TABLE width="100%" border=1>
    	    <TR>
        	  <TD noWrap align=center bgColor="#c0c0c0">
			  <a href="LoadMapForm.asp">Load Map</A></TD></TR></TABLE>
    	  <TABLE width="100%" border=1>
        	<TR>
	          <TD noWrap align=center bgColor="#c0c0c0">
			  <a href="UnloadMapForm.asp">Unload Map</A></TD></TR></TABLE>
		<%IF% IS_MAPS_DISPLAYED_CONFIGURABLE>
   		  <TABLE width="100%" border=1>
        	<TR>
	          <TD noWrap align=center bgColor="#c0c0c0">
			  <a href="SaveMapForm.asp">Save Map</A></TD></TR></TABLE>
		<%ENDIF%>
		<TABLE width="100%" border=1>
    	<TR>
			<TD noWrap align=center bgColor="#c0c0c0">
			<A href="notification.asp">Notifications</A></TD></TR></TABLE>
	<%ENDIF%>
<%ENDIF%>


<%IF% IS_CONFIGURE_REPORTS_ENABLED> 
			  <TABLE width="100%" border=1>
    			<TR>
        		  <TD noWrap align=center bgColor="#c0c0c0">
				  <A href="rcnotifylist.asp">Recurring Notifications</A></TD></TR></TABLE>

			<%IF% IS_USER_ACCESS_LOG> 
				<TABLE width="100%" border=1>
    				<TR>
        			<TD noWrap align=center bgColor="#c0c0c0">
					<A href="ReportPerformanceGraphs.asp">Performance Graphs</A></TD></TR></TABLE>
			<%ENDIF%>
<%ENDIF%>		
<%IF% IS_USER_CONFIGURE_USERS> 	 
    	  <TABLE width="100%" border=1>
        	<TR>
	          <TD noWrap align=center bgColor="#c0c0c0">
			  <A href="userlist.asp">Users</A></TD></TR></TABLE>
<%ENDIF%>
<%IF% IS_USER_ACCESS_LOG> 	 
      <TABLE width="100%" border=1>
        <TR>
          <TD noWrap align=center bgColor="#c0c0c0">
		  <A href="logview.asp">Log view</A></TD></TR></TABLE>
<%ENDIF%>

<%IF% IS_USER_ACCESS_TOOLS> 
	      <TABLE width="100%" border=1>
    	    <TR>
        	  <TD noWrap align=center bgColor="#c0c0c0">
			  <A href="tools.asp">Tools</A></TD></TR></TABLE>
<%ENDIF%>
	</TD></TR></TABLE>


<%INCLUDE% MainPageSuffix.asp>

<%include% StandardPageFooter.asp>
</BODY></HTML>