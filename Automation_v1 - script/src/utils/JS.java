package utils;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;

public class JS {

	public void checkPageIsReady(InternetExplorerDriver dr,Integer sec)
	{
		  
		  JavascriptExecutor js = (JavascriptExecutor)dr;

		  for (int i=0; i<5; i++)
		  { 
		   try 
		   {
		    Thread.sleep(1000*sec);
		   }catch (InterruptedException e) 
		   {
		    	System.out.println("Thread interrupt:"+Thread.interrupted());
		    	Thread.currentThread().interrupt(); //http://www.informit.com/articles/article.aspx?p=26326&seqNum=3
		   } 
		   //To check page ready state.
		   if (js.executeScript("return document.readyState").toString().equals("complete"))
		   { 
		    break; 
		   }
		  }
	}
	
	public void keyPress(InternetExplorerDriver dr,WebElement el, String ev, Integer kyCd, Integer frqncy) 
	{	
		JavascriptExecutor js = (JavascriptExecutor)dr;
		js.executeScript("function triggerEvent(el, type, keyCode) {\r\n" + 
				"    if ('createEvent' in document) {\r\n" + 
				"            // modern browsers, IE9+\r\n" + 
				"            var element = document.createEvent('HTMLEvents');\r\n" + 
				"            element.keyCode = keyCode;\r\n" + 
				"			element.bubbles = true;\r\n" + 
				"			element.eventPhase = 3;\r\n" + 
				"            element.initEvent(type, true, false);\r\n" + 
				"            el.dispatchEvent(element);\r\n" + 
				"    } else {\r\n" + 
				"        // IE 8\r\n" + 
				"        var e = document.createEventObject();\r\n" + 
				"        e.keyCode = keyCode;\r\n" + 
				"        e.eventType = type;\r\n" + 
				"        el.fireEvent('on'+e.eventType, e);\r\n" + 
				"    }\r\n" + 
				"}\r\n" + 
				"var el = arguments[0];\r\n"+
				"var ev = arguments[1];\r\n"+
				"var kyCd = arguments[2];\r\n"+
				"var frqncy = arguments[3];\r\n"+
				"\r\n"+
				"for(var i = 0; i<frqncy;i++){\r\n"+
				"triggerEvent(el, ev, kyCd);}", 
				el, ev, kyCd, frqncy);
	}
	
	public void mousePress(InternetExplorerDriver dr,WebElement el, String ev) 
	{
		
		JavascriptExecutor js = (JavascriptExecutor)dr;
		js.executeScript("function triggerMouseEvent(el, type) {\r\n" + 
				"    if ('createEvent' in document) {\r\n" + 
				"        // modern browsers, IE9+\r\n" + 
				"		var mouseEvent = document.createEvent ('MouseEvents');\r\n" + 
				"		mouseEvent.initEvent (type, true, true);\r\n" + 
				"		el.dispatchEvent (mouseEvent);\r\n" + 
				"    } else {\r\n" + 
				"        // IE 8\r\n" + 
				"		var mouseEvent = document.createEventObject();\r\n" + 
				"		mouseEvent.eventType = type;\r\n" + 
				"		el.fireEvent('on'+mouseEvent.eventType, mouseEvent);\r\n" + 
				"    }" + 
				"}\r\n"+
				"var el = arguments[0];\r\n"+
				"var ev = arguments[1];\r\n"+
				"triggerMouseEvent(el,ev);"
				, el, ev);
	}
	
	public void simulateMousePress(InternetExplorerDriver dr,WebElement el, String ev) {
		
		JavascriptExecutor js = (JavascriptExecutor)dr;
		
		js.executeScript("function simulateMouseEvent(element, eventName)\r\n" + 
				"{\r\n" + 
				"    var options = extend(defaultOptions, arguments[2] || {});\r\n" + 
				"    var oEvent, eventType = null;\r\n" + 
				"\r\n" + 
				"    for (var name in eventMatchers)\r\n" + 
				"    {\r\n" + 
				"        if (eventMatchers[name].test(eventName)) { eventType = name; break; }\r\n" + 
				"    }\r\n" + 
				"\r\n" + 
				"    if (!eventType)\r\n" + 
				"        throw new SyntaxError('Only HTMLEvents and MouseEvents interfaces are supported');\r\n" + 
				"\r\n" + 
				"    if (document.createEvent)\r\n" + 
				"    {\r\n" + 
				"        oEvent = document.createEvent(eventType);\r\n" + 
				"        if (eventType == 'HTMLEvents')\r\n" + 
				"        {\r\n" + 
				"            oEvent.initEvent(eventName, options.bubbles, options.cancelable);\r\n" + 
				"        }\r\n" + 
				"        else\r\n" + 
				"        {\r\n" + 
				"            oEvent.initMouseEvent(eventName, options.bubbles, options.cancelable, document.defaultView,\r\n" + 
				"            options.button, options.pointerX, options.pointerY, options.pointerX, options.pointerY,\r\n" + 
				"            options.ctrlKey, options.altKey, options.shiftKey, options.metaKey, options.button, element);\r\n" + 
				"        }\r\n" + 
				"        element.dispatchEvent(oEvent);\r\n" + 
				"    }\r\n" + 
				"    else\r\n" + 
				"    {\r\n" + 
				"        options.clientX = options.pointerX;\r\n" + 
				"        options.clientY = options.pointerY;\r\n" + 
				"        var evt = document.createEventObject();\r\n" + 
				"        oEvent = extend(evt, options);\r\n" + 
				"        element.fireEvent('on' + eventName, oEvent);\r\n" + 
				"    }\r\n" + 
				"    return element;\r\n" + 
				"}\r\n" + 
				"\r\n" + 
				"function extend(destination, source) {\r\n" + 
				"    for (var property in source)\r\n" + 
				"      destination[property] = source[property];\r\n" + 
				"    return destination;\r\n" + 
				"}\r\n" + 
				"\r\n" + 
				"var eventMatchers = {\r\n" + 
				"    'HTMLEvents': /^(?:load|unload|abort|error|select|change|submit|reset|focus|blur|resize|scroll)$/,\r\n" + 
				"    'MouseEvents': /^(?:click|dblclick|mouse(?:down|up|over|move|out))$/\r\n" + 
				"}\r\n" + 
				"var defaultOptions = {\r\n" + 
				"    pointerX: 0,\r\n" + 
				"    pointerY: 0,\r\n" + 
				"    button: 0,\r\n" + 
				"    ctrlKey: false,\r\n" + 
				"    altKey: false,\r\n" + 
				"    shiftKey: false,\r\n" + 
				"    metaKey: false,\r\n" + 
				"    bubbles: true,\r\n" + 
				"    cancelable: true\r\n" + 
				"}\r\n" + 
				"\r\n" +  
				"var el = arguments[0];\r\n"+
				"var ev = arguments[1];\r\n"+
				"simulateMouseEvent(el,ev);"
				, el, ev);
	}

}
