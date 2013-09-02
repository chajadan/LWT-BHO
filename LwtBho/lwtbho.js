function totalLeftOffset(element)
{
	var amount = 0;
	while (element != null)
	{
		amount += element.offsetLeft;
		element = element.offsetParent;
	}
	return amount;
}

function totalTopOffset(element)
{
	var amount = 0;
	while (element != null)
	{
		amount += element.offsetTop;
		element = element.offsetParent;
	}
	return amount;
}

function fixPageXY(e)
{
	if (e.pageX == null && e.clientX != null )
	{
		var html = document.documentElement;
		var body = document.body;
		e.pageX = e.clientX + (html.scrollLeft || body && body.scrollLeft || 0);
		e.pageX -= html.clientLeft || 0;
		e.pageY = e.clientY + (html.scrollTop || body && body.scrollTop || 0);
		e.pageY -= html.clientTop || 0;
	}
}

function setSelection(elem, lastterm, curTerm)
{
	elem.id = 'lwtcursel';
	lastterm.setAttribute('lwtcursel', curTerm);
	elem.className = elem.className + ' lwtSel';
}
function getSelection()
{
	var elem = document.getElementById('lwtcursel');
	if (elem === null)
		return '';
	return elem.getAttribute('lwtcursel');
}

function removeSelection()
{
	var elem = document.getElementById('lwtcursel');
	if (elem == null)
	{
		return;
	}
	elem.id = '';
	document.getElementById('lwtlasthovered').setAttribute('lwtcursel','');
	var curClass = elem.className;
	var lastSpace = curClass.lastIndexOf(' ');
	if (lastSpace >= 0)
	{
		var putClass = curClass.substr(0, lastSpace + 1);
		elem.className = putClass;
	}
}

function XOnElement(x, elem)
{
	if (x < totalLeftOffset(elem) || x >= totalLeftOffset(elem) + elem.offsetWidth)
	{
		return false;
	}
	return true;
}

function YOnElement(y, elem)
{
	if (y < totalTopOffset(elem) || y >= totalTopOffset(elem) + elem.offsetHeight)
	{
		return false;
	}
	return true;
}

function XYOnElement(x, y, elem)
{
	return (XOnElement(x, elem) && YOnElement(y, elem));
}

function lwtcontextexit()
{
	removeSelection();
	CloseOpenDialogs();
}

function lwtdivmout(e)
{
	fixPageXY(e);
	var statbox = window.curStatbox;
	var bVal = XYOnElement(e.pageX, e.pageY, statbox);
	if (bVal == false)
	{
		lwtmout(e);
	}
}

function lwtmout(e)
{
	fixPageXY(e);
	var statbox = window.curStatbox;
	if (!XYOnElement(e.pageX, e.pageY, statbox))
	{
		lwtcontextexit();
	}
}

function CloseOpenDialogs()
{
	window.curStatbox.style.display = 'none';
}

function lwtkeypress()
{
	if (document.getElementById('lwtlasthovered').getAttribute('lwtcursel') == '')
	{
		return;
	}
	var e = window.event;
	if (e.ctrlKey || e.altKey || e.shiftKey || e.metaKey)
	{
		return;
		alert('in lwtkeypress return!');
	}
	switch(e.keyCode)
	{
		case 49:
			document.getElementById('lwtsetstat1').click();
			break;
		case 50:
			document.getElementById('lwtsetstat2').click();
			break;
		case 51:
			document.getElementById('lwtsetstat3').click();
			break;
		case 52:
			document.getElementById('lwtsetstat4').click();
			break;
		case 53:
			document.getElementById('lwtsetstat5').click();
			break;
		case 87:
			document.getElementById('lwtsetstat99').click();
			break;
		case 73:
			document.getElementById('lwtsetstat98').click();
			break;
		case 85:
			document.getElementById('lwtsetstat0').click();
			break;
		case 27:
			CloseOpenDialogs();
			break;
	}
}

function ExtrapLink(elem, linkForm, curTerm)
{
	if (elem == null) {return;}
	var extrap = 'javascript:void(0);';
	if (linkForm.indexOf('*') == 0)
	{
		extrap = linkForm.substr(1, linkForm.length);
		extrap = extrap.replace('###', curTerm);
		elem.setAttribute('target', '_blank');
	}
	else if (linkForm.indexOf('###') >= 0)
	{
		extrap = linkForm.replace('###', curTerm);
		elem.setAttribute('target', '_blank');
	}
	elem.setAttribute('href', extrap);
}

function lwtmover(whichid, e, origin)
{
	removeSelection();
	var divRec = document.getElementById(whichid);
	if (divRec == null) {alert('could not locate divRec');}
	var curTerm = divRec.getAttribute('lwtterm');
	var lastterm = document.getElementById('lwtlasthovered');
	if (lastterm == null) {alert('could not loccated lasthovered field');}
	lastterm.setAttribute('lwtterm', curTerm);
	lastterm.setAttribute('lwtstat', divRec.getAttribute('lwtstat'));
	setSelection(origin, lastterm, curTerm);
	fixPageXY(e);
	lwtshowinlinestat(e, curTerm, origin);
	var infobox = document.getElementById('lwtinfobox');
	if (infobox == null) {alert('could not locate infobox to render popup');}
	infobox.style.left = (e.pageX + 10) + 'px';
	infobox.style.top = e.pageY + 'px';
	document.getElementById('lwtinfoboxterm').innerText = curTerm;
	document.getElementById('lwtinfoboxtrans').innerText = divRec.getAttribute('lwttrans');
	document.getElementById('lwtinfoboxrom').innerText = divRec.getAttribute('lwtrom');
}

function lwtshowinlinestat(e, curTerm, origin)
{
	var statbox = null;
	var curStat = '';

	if (window.mwTermBegin != null)
	{
		statbox = document.getElementById('lwtInlineMWEndPopup');
		if (window.keydownEventAttached == true)
		{
			document.detachEvent('onkeydown', lwtkeypress);
			window.keydownEventAttached = false;
		}
	}
	else
	{
		statbox = document.getElementById('lwtinlinestat');
		curStat = document.getElementById('lwtcursel').className;
		ExtrapLink(document.getElementById('lwtextrapdict1'), document.getElementById('lwtdict1').getAttribute('src'), curTerm);
		ExtrapLink(document.getElementById('lwtextrapdict2'), document.getElementById('lwtdict2').getAttribute('src'), curTerm);
		ExtrapLink(document.getElementById('lwtextrapgoogletrans'), document.getElementById('lwtgoogletrans').getAttribute('src'), curTerm);
		if (window.keydownEventAttached != true)
		{
			document.attachEvent('onkeydown', lwtkeypress);
			window.keydownEventAttached = true;
		}
	}

	if (statbox == null)
	{
		alert('could not locate inline status change popup');
		return;
	}

	window.curStatbox = statbox;

	var posElem = origin;
	var altElem = e.srcElement;
	var inlineTop2 = (totalTopOffset(posElem) + posElem.offsetHeight - 2);
	var inlineTop = totalTopOffset(posElem);
	if (curStat.indexOf('9') >= 0)
	{
		inlineTop -= 2;
	}


	statbox.style.left = totalLeftOffset(posElem) + 'px';
	document.getElementById('lwtTermTrans').style.width = posElem.offsetWidth + 'px';
	document.getElementById('lwtTermTrans2').style.width = posElem.offsetWidth + 'px';
	statbox.style.top = inlineTop + 'px';

	var inlineBottom = inlineTop + statbox.offsetHeight;
	statbox.style.display = 'block';
	if (inlineBottom > window.pageYOffset + window.innerHeight)
		statbox.scrollIntoView(false);
}

function traverseDomTree_NextNodeByTagName(elem, aTagName)
{
	if (elem.hasChildNodes() == true)
	{
		if (elem.firstChild.tagName == aTagName)
			return elem.firstChild;
		else
			return traverseDomTree_NextNodeByTagName(elem.firstChild, aTagName);
	}

	var possNextElem = elem.nextSibling;
	if (possNextElem == null)
	{
		while (elem.parentNode.nextSibling == null)
		{
			elem = elem.parentNode;
			if (elem == null)
				return null;
		}

		if (elem.parentNode.nextSibling.tagName == aTagName)
			return elem.parentNode.nextSibling;
		else
			return traverseDomTree_NextNodeByTagName(elem.parentNode.nextSibling, aTagName);
	}

	if (possNextElem.tagName == aTagName)
		return possNextElem;
	else
		return traverseDomTree_NextNodeByTagName(possNextElem, aTagName);
}

function beginMW()
{
	window.mwTermBegin = document.getElementById('lwtcursel');
	lwtcontextexit();
}

function cancelMW()
{
	window.mwTermBegin = null;
	lwtcontextexit();
}

function captureMW(newStatus)
{
	var newTerm = getChosenMWTerm(document.getElementById('lwtcursel'));
	cancelMW();
	if (newTerm != '')
	{
		var lastterm = document.getElementById('lwtlasthovered');
		if (lastterm == null) {alert('could not locate lasthovered field');}
		lastterm.setAttribute('lwtterm', newTerm);
		lastterm.setAttribute('lwtstat', '0');

		switch(newStatus)
		{
			case 49:
				document.getElementById('lwtsetstat1').click();
				break;
			case 50:
				document.getElementById('lwtsetstat2').click();
				break;
			case 51:
				document.getElementById('lwtsetstat3').click();
				break;
			case 52:
				document.getElementById('lwtsetstat4').click();
				break;
			case 53:
				document.getElementById('lwtsetstat5').click();
				break;
			case 87:
				document.getElementById('lwtsetstat99').click();
				break;
			case 73:
				document.getElementById('lwtsetstat98').click();
				break;
		}
	}
	else
		alert('Not a valid composite term selection.');
}

function getChosenMWTerm(elemLastPart)
{
	var bWithSpaces = false;
	if (document.getElementById('lwtwithspaces').getAttribute('value') == 'yes')
		bWithSpaces = true;

	var curTerm = window.mwTermBegin.getAttribute('lwtterm');
	if (curTerm.length <= 0)
		return '';
	var chosenMWTerm = curTerm;

	var elem = window.mwTermBegin;
	for (var i = 0; i < 8; i++)
	{
		elem = traverseDomTree_NextNodeByTagName(elem, 'SPAN');

		if (elem == null)
			return '';

		curTerm = elem.getAttribute('lwtterm');
		if (curTerm.length <= 0)
			i--;
		else
		{
			if (bWithSpaces == true)
				chosenMWTerm += ' ';
			
			chosenMWTerm += curTerm;
			alert(chosenMWTerm);

			if (elem == elemLastPart)
				return chosenMWTerm;
		}
	}

	return '';
}

function getPossibleMWTermParts(elem)
{
	var mwList = new Array(8);
	for (var i = 0; i < 8; i++)
	{
		mwList[i] = '';
	}

	var curTerm = elem.getAttribute('lwtterm');
		
	if (curTerm.length <= 0)
		return mwList;

	for (0; i < 8; i++)
	{
		elem = traverseDomTree_NextNodeByTagName(elem, 'SPAN');

		if (elem == null)
			return mwList;

		if (elem.className == 'lwtsent')
			return mwList;

		curTerm = elem.getAttribute('lwtterm');
		if (curTerm.length <= 0)
			i--;
		else
			mwList[i] = curTerm;
	}

	return mwList;
}
function getMWTermPartsAccrued(elem)
{
	var parts = getMWTermParts(elem);
	var partsAccrued = new Array(8);

	for (var i = 0; i < 8; i++)
	{
		mwList[i] = '';
	}

	var bWithSpaces = false;
	if (document.getElementById('lwtwithspaces').getAttribute('value') == 'yes')
		bWithSpaces = true;

	var curTerm = elem.getAttribute('lwtterm');
		
	if (curTerm.length <= 0)
		return mwList;

	var accruedTerm = curTerm;

	for (i = 0; i < 8 && parts[i] != ''; i++)
	{
		if (bWithSpaces == true)
			accruedTerm += ' ';

		accruedTerm += parts[i];
		partsAccrued[i] = accruedTerm;
	}

	return partsAccrued;
	
}