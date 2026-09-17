<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en">

<!-- #BeginTemplate "../jkp-ads_aspx.dwt" -->

<head>
<meta content="width=device-width, initial-scale=1" name="viewport" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<!-- #BeginEditable "doctitle" -->
<title>Formatted VBA ListBox control for Excel, Word and Access</title>
<!-- #EndEditable -->
<!-- #BeginEditable "headsection" -->
<meta content="An all-VBA formatted ListBox control for MSForms UserForms, with per-character font formatting, columns, keyboard control and configurable multi-select behavior." name="description" />
<meta content="VBA ListBox control, formatted VBA ListBox, Excel ListBox, Word ListBox, Access ListBox, MSForms ListBox, character formatting, multi-select ListBox, all VBA control, 64 bit Office" name="keywords" />
<meta content="en-us" http-equiv="Content-Language" />
<!-- #EndEditable --><% Response.Write(SiteTools.GetArticlejson()) %>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Organization",
  "name": "JKP Application Development Services",
  "url": "https://jkp-ads.com/",
  "logo": "https://jkp-ads.com/images/jkp-adslogo.gif",
  "founder": {
    "@type": "Person",
    "name": "Jan Karel Pieterse",
    "url": "https://jkp-ads.com/about.aspx"
  },
  "contactPoint": {
    "@type": "ContactPoint",
    "email": "info@jkp-ads.com",
    "contactType": "customer support",
    "availableLanguage": ["en", "nl"]
  },
  "sameAs": ["https://www.linkedin.com/in/jankarelpieterse"]
}
</script>
<!-- #BeginEditable "scripting" -->
<!-- Scripting area here -->
<!-- #EndEditable -->
<script src="/includes/topnav_script.js" type="text/javascript"></script>
<meta content="General" name="rating" />
<meta content="no" http-equiv="imagetoolbar" />
<meta content="Copyright ©, Jan Karel Pieterse All Rights Reserved" name="copyright" />
<meta content="nocache" name="robots" />
<meta content="noarchive" name="robots" />
<link href="../jkp-ads.css" rel="stylesheet" type="text/css" />
<% Response.Write(SiteTools.GetCanonical()) %>
</head>

<body id="top">
<div id="container">
	<div id="banner">
		<div id="logoleft">
			<a href="../index.aspx"><img alt="Home" src="../images/jkp-adslogo.gif" /></a>
		</div>
		<div id="logomiddle">
			<div class="search-container" role="search">
				<form action="https://www.google.com/search" method="get">
					<input name="ie" type="hidden" value="UTF-8" />
					<input name="oe" type="hidden" value="UTF-8" />
					<div class="smtxt"><label for="googlesearch">Site search</label></div>
					<input id="googlesearch" maxlength="255" name="q" placeholder="Search.." size="12" type="text" value="" />
					<button class="dropbtn" name="btnG" type="submit" value="?">?</button>
					<input name="domains" type="hidden" value="https://jkp-ads.com" />
					<input hidden="true" name="sitesearch" type="radio" value="" />
					<input checked="checked" hidden="true" name="sitesearch" type="radio" value="https://jkp-ads.com" />
				</form>
			</div>
		</div>
	</div>
	<% Response.write(SiteTools.GetNav()) %><br style="clear: both" />
	<% Response.write(SiteTools.GetCrumbs()) %><br />
	<div id="content" role="main">
		<%
If FeatureToggle.IsEnabled("AnnouncementEnabled") Then
    Dim lang = FeatureToggle.GetCurrentLanguage(HttpContext.Current)
    Dim heading = FeatureToggle.GetMessage("AnnouncementHeading", lang)
    Dim body = FeatureToggle.GetMessage("AnnouncementBody", lang)
%>
		<div style="border: 1px solid #ccc; background-color: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #d9534f;">
			<h2><%= HttpUtility.HtmlEncode(heading) %></h2>
			<p><%= HttpUtility.HtmlEncode(body) %></p>
		</div>
		<% End If %>
		<!-- #BeginEditable "content" -->
		<h1>An all-VBA formatted ListBox control for Excel, Word and Access</h1>
		<h2 id="Content0">Content</h2>
		<ul>
			<li><a href="#Introduction">Introduction</a></li>
			<li><a href="#Features">ListBox features</a></li>
			<li><a href="#Files">Files to import</a></li>
			<li><a href="#Setup">Add the control to a UserForm</a></li>
			<li><a href="#HowItWorks">How character formatting works</a></li>
			<li><a href="#Columns">Columns and bound values</a></li>
			<li><a href="#Selection">Selection and keyboard control</a></li>
			<li><a href="#Lifecycle">Resizing and cleanup</a></li>
			<li><a href="#Compatibility">Compatibility</a></li>
			<li><a href="#Disclaimer">Disclaimer</a></li>
			<li><a href="#faq">Frequently asked questions</a></li>
		</ul>
		<h2 id="Introduction">Introduction</h2>
		<p>The built-in MSForms ListBox is fast and convenient, but it you cannot 
		have different font or color settings for individual characters in a list 
		item. This all-VBA ListBox-style control gives you that capability while 
		retaining familiar properties such as <code>ListIndex</code>, <code>List</code>,
		<code>Selected</code>, <code>ColumnCount</code>, <code>BoundColumn</code> 
		and <code>MultiSelect</code>.</p>
		<p>The control is built within a standard MSForms Frame control placed on 
		a UserForm. It uses standard MSForms labels and text boxes at run time.</p>
		<p>Here is a screen-shot of a small dialog in my RefTreeAnalyser add-in 
		which uses this new control:</p>
		<p>
		<img alt="Screenshot of the Find a sheet dialog of RefTreeAnalyser, the ultimate Excel formula auditing tool." src="/images/vba-listbox-control01.png" /></p>
		<p>
		Notice how the "Name contains" characters are bolded and have a red font 
		in each worksheet's name.</p>
		<h2 id="Features">ListBox features</h2>
		<ul>
			<li>Per character you can set font name, size, bold, italic, underline, 
			strikethrough, weight and charset.</li>
			<li>Per character you can set foreground and background colors.</li>
			<li>It has Single, Multi and Extended selection modes.</li>
			<li>Supports Ctrl-click and Shift-click range selection.</li>
			<li>Supports Arrow, Home, End, Page Up, Page Down, Space and Enter keyboard 
			handling.</li>
			<li>It has multiple columns, zero-width hidden columns and bound values.</li>
			<li>Features automatic horizontal and vertical Frame scrollbars.</li>
			<li>Explicit cleanup of all dynamically created controls.</li>
		</ul>
		<h2 id="Files">Files to import</h2>
		<p>The complete reusable control consists of three class modules:</p>
		<ul>
			<li><code>clsFormattedListBox.cls</code>: The main class, which is the 
			controller offering a familiar ListBox API.</li>
			<li><code>clsFormattedListItem.cls</code>: The class that handles each 
			list item. It contains one row of values, and handles formatting and 
			run-time controls.</li>
			<li><code>clsFormattedCharacters.cls</code>: The class that handles 
			the character-range formatting.</li>
		</ul>
		<h2 id="Setup">Add the control to a UserForm</h2>
		<p>Add an MSForms Frame named <code>frmList</code> (or use any other convenient 
		control name of course) to your UserForm. Leave its caption empty and set 
		its initial ScrollBars property to None. Set its Font to the font you want 
		to serve as default for all controls in the list box. Then add the following 
		code to the form:</p>
		<div class="vbacode">
			<span class="kwrd">Option</span> <span class="kwrd">Explicit</span><br />
			<br />
			<span class="kwrd">Private</span> <span class="kwrd">WithEvents</span> 
			mList <span class="kwrd">As</span> clsFormattedListBox<br />
			<br />
			<span class="kwrd">Private</span> <span class="kwrd">Sub</span> UserForm_Initialize()<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> mList = <span class="kwrd">
			New</span> clsFormattedListBox<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">With</span> mList<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.BeginUpdate<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> .HostFrame 
			= Me.frmList<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ColumnCount = 2<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.ColumnWidths = "180;0"<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.BoundColumn = 2<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.MultiSelect = fmMultiSelectExtended<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.AddItem "First item"<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.List(0, 1) = "item-key-1"<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;.EndUpdate<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">End</span> <span class="kwrd">With</span><br />
			End <span class="kwrd">Sub</span><br />
		</div>
		<p>Of course the lines starting with .AddItem and .List(0, 1) are there 
		for demonstration only, I leave it up to you to replace them with relevant 
		code that adds the list.</p>
		<h2 id="HowItWorks">How character formatting works</h2>
		<p>Each logical list item is represented by one <code>clsFormattedListItem</code> 
		instance. A visible cell is divided into contiguous characters that share 
		the same formatting. Each set of characters is drawn as a separate label.</p>
		<p>The control applies the character set's font settings, sets the label's
		<code>AutoSize</code> property to true, reads its resulting width and places 
		the next set of characters directly beside it. This follows the same technique 
		as our <a href="treeview.aspx">all-VBA TreeView</a>.</p>
		<p><code>Characters</code> uses a 1-based start position and length, like 
		Excel's Characters object. Rows and columns use the native ListBox-style 
		0-based indexes. Here is some example code that deomstrates how you might 
		change the font characteristics of some characters of a list item:</p>
		<div class="vbacode">
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Dim</span> oItem <span class="kwrd">As</span> 
			clsFormattedListItem<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Dim</span> oCharacters
			<span class="kwrd">As</span> clsFormattedCharacters<br />
			<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> oItem = mList.AddItem("Sheet 
			1")<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> oCharacters = oItem.Characters(2, 
			1)<br />
&nbsp;&nbsp;&nbsp;&nbsp;oCharacters.Bold = <span class="kwrd">True</span><br />
&nbsp;&nbsp;&nbsp;&nbsp;oCharacters.ForeColor = RGB(192, 0, 0)<br />
			<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> oCharacters = oItem.Characters(5, 
			1)<br />
&nbsp;&nbsp;&nbsp;&nbsp;oCharacters.Bold = <span class="kwrd">True</span></div>
		<p>Adjacent characters with identical settings are rendered as one label.</p>
		<h2 id="Columns">Columns and bound values</h2>
		<p>Set <code>ColumnCount</code> before adding values. <code>ColumnWidths</code> 
		is a semicolon-separated list measured in points; use zero to hide a data 
		column. <code>BoundColumn</code> is 1-based, matching MSForms. A BoundColumn 
		of zero returns the row index instead of the content of one of the list 
		row columns.</p>
		<div class="vbacode">
&nbsp;&nbsp;&nbsp;&nbsp;mList.ColumnCount = 2<br />
&nbsp;&nbsp;&nbsp;&nbsp;mList.ColumnWidths = "180;0"<br />
&nbsp;&nbsp;&nbsp;&nbsp;mList.BoundColumn = 2<br />
&nbsp;&nbsp;&nbsp;&nbsp;mList.List(0, 1) = "hidden-value"<br />
&nbsp;&nbsp;&nbsp;&nbsp;Debug.Print mList.Value</div>
		<h2 id="Selection">Selection and keyboard control</h2>
		<p>The <code>MultiSelect</code> property accepts the built-in MSForms constants
		<code>fmMultiSelectSingle</code>, <code>fmMultiSelectMulti</code> and
		<code>fmMultiSelectExtended</code>. Use <code>Selected(row)</code> to inspect 
		or change an item's selection state.</p>
		<div class="vbacode">
			<span class="kwrd">Private</span> <span class="kwrd">Sub</span> mList_Click(<span class="kwrd">ByVal</span> 
			Index <span class="kwrd">As</span> <span class="kwrd">Long</span>)<br />
&nbsp;&nbsp;&nbsp;&nbsp;Debug.Print Index, mList.Value<br />
			<span class="kwrd">End</span> <span class="kwrd">Sub</span><br />
			<br />
			<span class="kwrd">Private</span> <span class="kwrd">Sub</span> mList_DblClick(<span class="kwrd">ByVal</span> 
			Index <span class="kwrd">As</span> <span class="kwrd">Long</span>)<br />
&nbsp;&nbsp;&nbsp;&nbsp;Unload Me<br />
			<span class="kwrd">End</span> <span class="kwrd">Sub</span></div>
		<p>The control also raises <code>Change</code>, <code>KeyDown</code> and
		<code>KeyUp</code> events. The focused row is kept visible while the user 
		moves through the list.</p>
		<h2 id="Lifecycle">Resizing and cleanup</h2>
		<p>Call <code>Resize</code> after changing the host Frame's dimensions. 
		Call <code>Terminate</code> when the UserForm is destroyed so all generated 
		controls are properly removed.</p>
		<div class="vbacode">
			<span class="kwrd">Private</span> <span class="kwrd">Sub</span> UserForm_Resize()<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">If</span> <span class="kwrd">Not</span> 
			mList <span class="kwrd">Is</span> <span class="kwrd">Nothing</span>
			<span class="kwrd">Then</span> mList.Resize<br />
			<span class="kwrd">End</span> <span class="kwrd">Sub</span><br />
			<br />
			<span class="kwrd">Private</span> <span class="kwrd">Sub</span> UserForm_Terminate()<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">If</span> <span class="kwrd">Not</span> 
			mList <span class="kwrd">Is</span> <span class="kwrd">Nothing</span>
			<span class="kwrd">Then</span> mList.Terminate<br />
&nbsp;&nbsp;&nbsp;&nbsp;<span class="kwrd">Set</span> mList = <span class="kwrd">
			Nothing</span><br />
			<span class="kwrd">End</span> <span class="kwrd">Sub</span></div>
		<p>The Frame supports scrolling. Mouse-wheel behavior must be supplied by 
		your own project and is deliberately not included in the three-class bundle.</p>
		<h2 id="Compatibility">Compatibility</h2>
		<p>The control contains no Windows API declarations and no Excel object-model 
		dependency. It is designed for MSForms projects, including 32-bit and 64-bit 
		Office. It is up to you to check behavior in each Office host and Mac version 
		required by your own deployment before distribution.</p>
		<h2 id="Disclaimer">Disclaimer</h2>
		<p>You use this control at your own risk. JKP Application Development Services 
		accepts no liability for damages arising from its use. Test the control 
		thoroughly in every Office host and platform supported by your project.</p>
		<h2 id="faq">Frequently asked questions</h2>
		<div class="smtxt">
			<p><a href="#Introduction">Why use a custom formatted VBA ListBox?</a></p>
			<p><a href="#HowItWorks">How can separate characters use different formatting?</a></p>
			<p><a href="#Files">Which class modules must I import?</a></p>
			<p><a href="#Setup">How do I attach the control to a UserForm?</a></p>
			<p><a href="#Formatting">How do I make found characters bold?</a></p>
			<p><a href="#Selection">Which multi-select modes and keys are supported?</a></p>
			<p><a href="#Compatibility">Does the control use Windows API calls?</a></p>
		</div>
		<div id="comments">
			<hr />
			<h2>Comments</h2>
			<div commentsection="/includes/getcomments.aspx?page=articles/vba-listbox-control.aspx">
				Loading comments...</div>
			<div commentform="/includes/comments.aspx">
			</div>
		</div>
		<p>&nbsp;</p>
		<!-- #EndEditable -->
	</div>
	<div id="footer" role="contentinfo"><% Response.write(SiteTools.GetFooter()) %></div>
</div>
</body>

<!-- #EndTemplate -->
</html>
