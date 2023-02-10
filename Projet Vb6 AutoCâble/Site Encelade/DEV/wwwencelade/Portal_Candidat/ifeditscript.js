    function UtilBeginScript()
    {
	return String.fromCharCode(60, 115, 99, 114, 105, 112, 116, 62);
    }

    function UtilEndScript()
    {
	return String.fromCharCode(60, 47, 115, 99, 114, 105, 112, 116, 62);
    }

	function IDGenerator(nextID)
	{
		this.nextID = nextID;
		this.GenerateID = IDGeneratorGenerateID;
	}

	function IDGeneratorGenerateID()
	{
		return this.nextID++;
	}

	var BUTTON_IMAGE_PREFIX = "buttonImage";
	var BUTTON_DIV_PREFIX = "buttonDiv";
	var BUTTON_PAD1_PREFIX = "buttonPad1";
	var BUTTON_PAD2_PREFIX = "buttonPad2";
	var buttonMap = new Object();

	function Button
	(
		idGenerator,
		caption,
		action,
		image
	)
	{
		this.idGenerator = idGenerator;
		this.caption = caption;
		this.action = action;
		this.image = image;
		this.enabled = true;
		this.Instantiate = ButtonInstantiate;
		this.Enable = ButtonEnable;
	}

	function ButtonInstantiate()
	{
		this.id = this.idGenerator.GenerateID();
		buttonMap[this.id] = this;
		var html = "";
		html += '<div id="';
		html += BUTTON_DIV_PREFIX;
		html += this.id;
		html += '" class="ButtonNormal"';
		html += ' onselectstart="ButtonOnSelectStart()"';
		html += ' ondragstart="ButtonOnDragStart()"';
		html += ' onmousedown="ButtonOnMouseDown(this)"';
		html += ' onmouseup="ButtonOnMouseUp(this)"';
		html += ' onmouseout="ButtonOnMouseOut(this)"';
		html += ' onmouseover="ButtonOnMouseOver(this)"';
		html += ' onclick="ButtonOnClick(this)"';
		html += ' ondblclick="ButtonOnDblClick(this)"';
		html += '>';
		html += '<table cellpadding=0 cellspacing=0 border=0><tr><td><img id="';
		html += BUTTON_PAD1_PREFIX;
		html += this.id;
		html += '" width=2 height=2></td><td></td><td></td></tr><tr><td></td><td>';
		html += '<img id="';
		html += BUTTON_IMAGE_PREFIX;
		html += this.id;
		html += '" src="';
		html += this.image;
		html += '" title="';
		html += this.caption;
		html += '" class="Image"';
		html += '>';
		html += '</td><td></td></tr><tr><td></td><td></td><td><img id="';
		html += BUTTON_PAD2_PREFIX;
		html += this.id;
		html += '" width=2 height=2></td></tr></table>';
		html += '</div>';
		document.write(html);
	}

	function ButtonEnable(enabled)
	{
		this.enabled = enabled;
		if (this.enabled)
		{
			document.all[BUTTON_DIV_PREFIX + this.id].className = "ButtonNormal";
		}
		else
		{
			document.all[BUTTON_DIV_PREFIX + this.id].className = "ButtonDisabled";
		}
	}

	function ButtonOnSelectStart()
	{
		window.event.returnValue = false;
	}

	function ButtonOnDragStart()
	{
		window.event.returnValue = false;
	}

	function ButtonOnMouseDown(element)
	{
		if (event.button == 1)
		{
			var id = element.id.substring(BUTTON_DIV_PREFIX.length, element.id.length);
			var button = buttonMap[id];
			if (button.enabled)
			{
				ButtonPushButton(id);
			}
		}
	}

	function ButtonOnMouseUp(element)
	{
		if (event.button == 1)
		{
			var id = element.id.substring(BUTTON_DIV_PREFIX.length, element.id.length);
			var button = buttonMap[id];
			if (button.enabled)
			{
				ButtonReleaseButton(id);
			}
		}
	}

	function ButtonOnMouseOut(element)
	{
		var id = element.id.substring(BUTTON_DIV_PREFIX.length, element.id.length);
		var button = buttonMap[id];
		if (button.enabled)
		{
			ButtonReleaseButton(id);
		}
	}

	function ButtonOnMouseOver(element)
	{
		var id = element.id.substring(BUTTON_DIV_PREFIX.length, element.id.length);
		var button = buttonMap[id];
		if (button.enabled)
		{
			ButtonReleaseButton(id);
			document.all[BUTTON_DIV_PREFIX + id].className = "ButtonMouseOver";
		}
	}

	function ButtonOnClick(element)
	{
		var id = element.id.substring(BUTTON_DIV_PREFIX.length, element.id.length);
		var button = buttonMap[id];
		if (button.enabled)
		{
			eval(button.action);
		}
	}

	function ButtonOnDblClick(element)
	{
		ButtonOnClick(element);
	}

	function ButtonPushButton(id)
	{
		document.all[BUTTON_PAD1_PREFIX + id].width = 3;
		document.all[BUTTON_PAD1_PREFIX + id].height = 3;
		document.all[BUTTON_PAD2_PREFIX + id].width = 1;
		document.all[BUTTON_PAD2_PREFIX + id].height = 1;
		document.all[BUTTON_DIV_PREFIX + id].className = "ButtonPressed";
	}

	function ButtonReleaseButton(id)
	{
		document.all[BUTTON_PAD1_PREFIX + id].width = 2;
		document.all[BUTTON_PAD1_PREFIX + id].height = 2;
		document.all[BUTTON_PAD2_PREFIX + id].width = 2;
		document.all[BUTTON_PAD2_PREFIX + id].height = 2;
		document.all[BUTTON_DIV_PREFIX + id].className = "ButtonNormal";
	}






    var IMAGE_CHOOSER_DIV_PREFIX = "imageChooserDiv";
    var IMAGE_CHOOSER_IMG_PREFIX = "imageChooserImg";
    var imageChooserMap = new Object();

    function ImageChooser
    (
	    idGenerator,
	    numRows,
	    numCols,
	    images,
	    callback
    )
    {
	    this.idGenerator = idGenerator;
	    this.numRows = numRows;
	    this.numCols = numCols;
	    this.images = images;
	    this.callback = callback;
	    this.Instantiate = ImageChooserInstantiate;
	    this.Show = ImageChooserShow;
	    this.Hide = ImageChooserHide;
	    this.IsShowing = ImageChooserIsShowing;
	    this.SetUserData = ImageChooserSetUserData;
    }

    function ImageChooserInstantiate()
    {
	    this.id = this.idGenerator.GenerateID();
	    imageChooserMap[this.id] = this;
	    var html = '';
	    html += '<table>';
	    html += '<tr>';
	    html += '<td>';
	    html += '<div id="' + IMAGE_CHOOSER_DIV_PREFIX + this.id + '" style="display:none;position:absolute;background-color:buttonface;border-left:buttonhighlight solid 1px;border-top:buttonhighlight solid 1px;border-right:buttonshadow solid 1px;border-bottom:buttonshadow solid 1px">';
	    html += '<table>';
	    for (var i = 0; i < this.numRows; i++) {
		    html += '<tr>';
		    for (var j = 0; j < this.numCols; j++) {
			    html += '<td>';
			    var k = i * this.numCols + j;
				    html += '<img id="' + IMAGE_CHOOSER_IMG_PREFIX + this.id + '_' + k + '" src="' + this.images[k] + '" onmouseover="ImageChooserOnMouseOver()" onmouseout="ImageChooserOnMouseOut()" onclick="ImageChooserOnClick()">';
			    html += '</td>';
		    }
		    html += '</tr>';
	    }
	    html += '</table>';
	    html += '</div>';
	    html += '</td>';
	    html += '</tr>';
	    html += '</table>';
	    document.write(html);
    }

    function ImageChooserShow(x, y)
    {
	    eval(IMAGE_CHOOSER_DIV_PREFIX + this.id).style.left = x;
	    eval(IMAGE_CHOOSER_DIV_PREFIX + this.id).style.top = y;
	    eval(IMAGE_CHOOSER_DIV_PREFIX + this.id).style.display = "block";
    }

    function ImageChooserHide()
    {
	    eval(IMAGE_CHOOSER_DIV_PREFIX + this.id).style.display = "none";
    }

    function ImageChooserIsShowing()
    {
	    return eval(IMAGE_CHOOSER_DIV_PREFIX + this.id).style.display == "block";
    }

    function ImageChooserSetUserData(userData)
    {
	this.userData = userData;
    }

    function ImageChooserOnMouseOver()
    {
	    if (event.srcElement.tagName == "IMG") {
		    event.srcElement.style.backgroundColor = "highlight";
	    }
    }

    function ImageChooserOnMouseOut()
    {
	    if (event.srcElement.tagName == "IMG") {
		    event.srcElement.style.backgroundColor = "buttonface";
	    }
    }

    function ImageChooserOnClick()
    {
	    if (event.srcElement.tagName == "IMG") {
		    var underscore = event.srcElement.id.indexOf("_");
		    if (underscore != -1) {
			    var id = event.srcElement.id.substring(IMAGE_CHOOSER_IMG_PREFIX.length, underscore);
			    var imageChooser = imageChooserMap[id];
			    imageChooser.Hide();
			    var index = event.srcElement.id.substring(underscore + 1);
			    if (imageChooser.callback) {
				    imageChooser.callback(imageChooser.images[index], imageChooser.userData);
			    }
		    }
	    }
    }

	var EDITOR_COMPOSITION_PREFIX = "editorComposition";
	var EDITOR_PARAGRAPH_PREFIX = "editorParagraph";
	var EDITOR_LIST_AND_INDENT_PREFIX = "editorListAndIndent";
	var EDITOR_TOP_TOOLBAR_PREFIX = "editorTopToolbar";
	var EDITOR_BOTTOM_TOOLBAR_PREFIX = "editorBottomToolbar";
	var editorMap = new Object();
	var editorIDGenerator = null;

	function Editor(idGenerator)
	{
		this.idGenerator = idGenerator;
		this.textMode = false;
		this.brief = false;
		this.instantiated = false;
		this.Instantiate = EditorInstantiate;
		this.GetText = EditorGetText;
		this.SetText = EditorSetText;
		this.GetHTML = EditorGetHTML;
		this.SetHTML = EditorSetHTML;
		this.GetBrief = EditorGetBrief;
		this.SetBrief = EditorSetBrief;
	}

	function EditorInstantiate()
	{
		if (this.instantiated) {
			return;
		}
		this.id = this.idGenerator.GenerateID();
		editorMap[this.id] = this;
		editorIDGenerator = this.idGenerator;

		var html = '';
		html += '<table cellpadding="0" cellspacing="0" border="0" width="100%" height="95%">';
		html += '<tr>';
		html += '<td id="' + EDITOR_TOP_TOOLBAR_PREFIX + this.id + '" class="Toolbar">';
		html += '<table cellpaddin="0" cellspacing="0" border="0">';
		html += '<tr>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += '<select class="List" onchange="EditorOnFont(' + this.id + ', this)">';
		html += '<option class="Heading">Font</option>';
		html += '<option value="Arial">Arial</option>';
		html += '<option value="Arial Black">Arial Black</option>';
		html += '<option value="Arial Narrow">Arial Narrow</option>';
		html += '<option value="Comic Sans MS">Comic Sans MS</option>';
		html += '<option value="Courier New">Courier New</option>';
		html += '<option value="System">System</option>';
		html += '<option value="Times New Roman">Times New Roman</option>';
		html += '<option value="Verdana">Verdana</option>';
		html += '<option value="Wingdings">Wingdings</option>';
		html += '</select>';
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += '<select class="List" onchange="EditorOnSize(' + this.id + ', this)">';
		html += '<option class="Heading">Size</option>';
		html += '<option value="1">1</option>';
		html += '<option value="2">2</option>';
		html += '<option value="3">3</option>';
		html += '<option value="4">4</option>';
		html += '<option value="5">5</option>';
		html += '<option value="6">6</option>';
		html += '<option value="7">7</option>';
		html += '</select>';
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += '<select class="List" onchange="EditorOnJoeColor(' + this.id + ', this, this.value);this.selectedIndex=0;">';
		html += '<option class="Heading">Color</option>';
		html += '<option style="color: black; background-color: black;" value="000000">black</option>';
		html += '<option style="color: blue; background-color: blue;" value="0000FF">blue</option>';
		html += '<option style="color: red; background-color: red;" value="FF0000">red</option>';
		html += '<option style="color: green; background-color: green;" value="green">green</option>';
		html += '<option style="color: purple; background-color: purple;" value="purple">purple</option>';
		html += '<option style="color: yellow; background-color: yellow;" value="yellow">yellow</option>';
		html += '<option style="color: orange; background-color: orange;" value="orange">orange</option>';
		html += '<option style="color: lime; background-color: lime;" value="lime">lime</option>';
		html += '<option style="color: aqua; background-color: aqua;" value="aqua">aqua</option>';
		html += '<option style="color: white; background-color: white;" value="white">white</option>';
		html += '</select>';
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var boldButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Bold",';
		html += '"EditorOnBold(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/bold.gif"';
		html += ');';
		html += 'boldButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var italicButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Italic",';
		html += '"EditorOnItalic(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/italic.gif"';
		html += ');';
		html += 'italicButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var underlineButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Underline",';
		html += '"EditorOnUnderline(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/under.gif"';
		html += ');';
		html += 'underlineButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var alignLeftButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Align Left",';
		html += '"EditorOnAlignLeft(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/leftalign.gif"';
		html += ');';
		html += 'alignLeftButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var centerButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Center",';
		html += '"EditorOnCenter(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/centeralign.gif"';
		html += ');';
		html += 'centerButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var alignRightButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Align Right",';
		html += '"EditorOnAlignRight(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/rightalign.gif"';
		html += ');';
		html += 'alignRightButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += '<div class="Space"></div>';
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var createHyperlinkButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Create Hyperlink",';
		html += '"EditorOnCreateHyperlink(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/createlink.gif"';
		html += ');';
		html += 'createHyperlinkButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';
		html += UtilBeginScript();
		html += 'var createImglinkButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Insert Image",';
		html += '"EditorOnCreateImage(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/createimage.gif"';
		html += ');';
		html += 'createImglinkButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';	
		html += UtilBeginScript();
		html += 'var createTextlinkButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Insert Text or Html File",';
		html += '"EditorOnCreateText(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/createText.gif"';
		html += ');';
		html += 'createTextlinkButton.Instantiate();';
		html += UtilEndScript();
		html += '</td>';
		html += '<td>';	
		html += UtilBeginScript();
		html += 'var createTableButton = new Button(';
		html += 'editorIDGenerator,';
		html += '"Insert Frame",';
		html += '"EditorOnCreateTable(' + this.id + ')",';
		html += '"../Portal_HTMLEditor/frame.gif"';
		html += ');';
		html += 'createTableButton.Instantiate();';
		html += UtilEndScript();		
		html += '</td>';
		html += '</tr>';
		html += '</table>';
		html += '</td>';
		html += '</tr>';
		html += '<tr>';
		html += '<td>';
		html += '<iframe src="" scrolling=auto id="' + EDITOR_COMPOSITION_PREFIX + this.id + '" width="100%" height="100%">';
		html += '</iframe>';
		html += '</td>';
		html += '</tr>';
		html += '</table>';
		document.write(html);

		html = '<link rel=STYLESHEET href="' + HtmlEditorStyle + '" type="text/css"><body></body>';
		eval(EDITOR_COMPOSITION_PREFIX + this.id).document.write(html);
		eval(EDITOR_COMPOSITION_PREFIX + this.id).document.designMode = "on";

		editorIDGenerator = null;
		this.instantiated = true;
	}

	function  EditorGetText()
	{
		return eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerText;
	}

	function  EditorSetText(text)
	{
		text = text.replace(/\n/g, "<br>");
		eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerHTML = text;
	}

	function  EditorGetHTML()
	{
		if (this.textMode) {
			return eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerText;
		}
		EditorCleanHTML(this.id);
		EditorCleanHTML(this.id);
		return eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerHTML;
	}

	function  EditorSetHTML(html)
	{
		if (this.textMode) {
			eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerText = html;
		}
		else {
			eval(EDITOR_COMPOSITION_PREFIX + this.id).document.body.innerHTML = html;
		}
	}
	
	function EditorGetBrief()
	{
		return this.brief;
	}

	function EditorSetBrief(brief)
	{
		this.brief = brief;
		var display = this.brief ? "none" : "inline";
		if (this.instantiated) {
			eval(EDITOR_PARAGRAPH_PREFIX + this.id).style.display = display;
			eval(EDITOR_LIST_AND_INDENT_PREFIX + this.id).style.display = display;
		}
	}

	function EditorOnCreateHyperlink(id)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		var anchor = EditorGetElement("A", eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange().parentElement());
		var link = prompt("Enter Link Location (eg. http://www.euxia.com):", anchor ? anchor.href : "http://");
		if (link && link != "http://") {
			if (eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.type == "None") {
				var selection = eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange();
				selection.pasteHTML('<A TARGET="_NEW" HREF="' + link + '"></A>');
				selection.select();
			}
			else {
				EditorFormat(id, "CreateLink", link);
			}
		}
	}
	
	function EditorOnCreateImage(id)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		var anchor = EditorGetElement("A", eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange().parentElement());
		window.open('../Portal_htmleditor/portal_upload.asp?mode=img&ID=' + id,'upload','resizable=1,status=no,top=250,left=300,width=500,height=300');
	}
	
	function EditorOnCreateText(id)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		var anchor = EditorGetElement("A", eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange().parentElement());
		window.open('../Portal_htmleditor/portal_upload.asp?mode=txt&ID=' + id,'upload','resizable=1,status=no,top=250,left=300,width=500,height=300');
	}
	
	function EditorOnCreateTable(id)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		var anchor = EditorGetElement("A", eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange().parentElement());
		window.open('../Portal_htmleditor/portal_Table.asp?ID=' + id,'upload','resizable=1,status=no,top=250,left=300,width=300,height=175');
	}
	
	function EditorOnCreateHyperlink3(id)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		var anchor = EditorGetElement("A", eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange().parentElement());
		var link = prompt("Enter Image Location (eg. http://www.euxia.com/logo.gif):", anchor ? anchor.href : "http://");
		if (link && link != "http://") {
			if (eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.type == "None") {
				var selection = eval(EDITOR_COMPOSITION_PREFIX + id).document.selection.createRange();
				selection.pasteHTML('<IMG SRC="' + link + '">');
				selection.select();
			}
			else {
				EditorFormat(id, "CreateLink", link);
			}
		}
	}

	function EditorOnAlignLeft(id)
	{
		EditorFormat(id, "justifyleft");
	}

	function EditorOnCenter(id)
	{
		EditorFormat(id, "justifycenter");
	}

	function EditorOnAlignRight(id)
	{
		EditorFormat(id, "justifyright");
	}
	
	function EditorOnCut(id)
	{
		EditorFormat(id, "cut");
	}

	function EditorOnCopy(id)
	{
		EditorFormat(id, "copy");
	}

	function EditorOnPaste(id)
	{
		EditorFormat(id, "paste");
	}

	function EditorOnBold(id)
	{
		EditorFormat(id, "bold");
	}

	function EditorOnItalic(id)
	{
		EditorFormat(id, "italic");
	}

	function EditorOnUnderline(id)
	{
		EditorFormat(id, "underline");
	}

	function EditorOnJoeColor(id, color, color2)
	{
		EditorFormat(id, "forecolor", color2);
	}

	function EditorOnFont(id, select)
	{
		EditorFormat(id, "fontname", select[select.selectedIndex].value);
		select.selectedIndex = 0;
	}

	function EditorOnSize(id, select)
	{
		EditorFormat(id, "fontsize", select[select.selectedIndex].value);
		select.selectedIndex = 0;
	}

	function EditorOnViewHTMLSource(id, textMode)
	{
		var editor = editorMap[id];
		editor.textMode = textMode;
		if (editor.textMode) {
			EditorCleanHTML(id);
			EditorCleanHTML(id);
			eval(EDITOR_COMPOSITION_PREFIX + id).document.body.innerText = eval(EDITOR_COMPOSITION_PREFIX + id).document.body.innerHTML;
		}
		else {
			eval(EDITOR_COMPOSITION_PREFIX + id).document.body.innerHTML = eval(EDITOR_COMPOSITION_PREFIX + id).document.body.innerText;
		}
		eval(EDITOR_COMPOSITION_PREFIX + id).focus();
	}
		
	function EditorValidateMode(id)
	{
		var editor = editorMap[id];
		if (!editor.textMode) {
			return true;
		}
		alert('Please uncheck the "View HTML Source" checkbox to use the toolbars.');
		eval(EDITOR_COMPOSITION_PREFIX + id).focus();
		return false;
	}

	function EditorFormat(id, what, opt)
	{
		if (!EditorValidateMode(id)) {
			return;
		}
		if (opt == "removeFormat") {
			what = opt;
			opt = null;
		}
		if (opt == null) {
			eval(EDITOR_COMPOSITION_PREFIX + id).document.execCommand(what);
		}
		else {
			eval(EDITOR_COMPOSITION_PREFIX + id).document.execCommand(what, "", opt);
		}
	}

	function EditorCleanHTML(id)
	{
		var fonts = eval(EDITOR_COMPOSITION_PREFIX + id).document.body.all.tags("FONT");
		for (var i = fonts.length - 1; i >= 0; i--) {
			var font = fonts[i];
			if (font.style.backgroundColor == "#ffffff") {
				font.outerHTML = font.innerHTML;
			}
		}
	}

	function EditorGetElement(tagName, start)
	{
		while (start && start.tagName != tagName) {
			start = start.parentElement;
		}
		return start;
	}


function Switch() {
    if (editor.GetText() != "" && editor.GetText() != editor.GetHTML()) {
    conf = confirm("This will convert your message into plain text.  All formatting will be lost.  Continue?");
    if (!conf) return;
  }
  document.Compose.Body.value = editor.GetText();
    document.Compose.action = document.Compose.action + "&SWITCH=1";
  document.Compose.submit();
}

//
//function document.body.onload() {
//  editor.SetHTML(document.all.plainmsg.innerHTML);
//}

function RecordAttachments(Files, File0, File1, File2) {
  window.document.Compose.elements["File0Data"].value = File0;
  window.document.Compose.elements["File1Data"].value = File1;
  window.document.Compose.elements["File2Data"].value = File2;
    window.document.all.Atts.innerText = Files;
  }

var remote=null;
function rs(n,u,w,h,x) {
  args="width="+w+",height="+h+",resizable=yes,scrollbars=yes,status=0";
  remote=window.open(u,n,args);
  if (remote != null) {
    if (remote.opener == null)
      remote.opener = self;
  }
  if (x == 1) { return remote; }
}

var awnd=null;
function ScriptAttach() {
  f0 = escape(document.Compose.elements["File0Data"].value);
  f1 = escape(document.Compose.elements["File1Data"].value);
  f2 = escape(document.Compose.elements["File2Data"].value);
  fname = escape(document.Compose.elements["FName"].value);
  awnd=rs('att','/ym/Attachments?YY=27590&File0Data='+f0+'&File1Data='+f1+'&File2Data='+f2+'&FName='+fname,450,600,1);
  awnd.focus();
}

var sigAttMap = [false];

function OnFromAddrChange()
{
    var i = document.Compose.FromAddr.selectedIndex;
    if (i >= 0 && i < sigAttMap.length) {
	document.all.SA.checked = sigAttMap[i];
    }
}

	var idGenerator = new IDGenerator(0);
	var editor = new Editor(idGenerator);
	editor.Instantiate();
