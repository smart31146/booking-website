  1 ?/*
  2 Copyright (c) 2003-2010, CKSource - Frederico Knabben. All rights reserved.
  3 For licensing, see LICENSE.html or http://ckeditor.com/license
  4 */
  5 
  6 /**
  7  * @fileOverview The "filebrowser" plugin, it adds support for file uploads and
  8  *               browsing.
  9  *
 10  * When file is selected inside of the file browser or uploaded, its url is
 11  * inserted automatically to a field, which is described in the 'filebrowser'
 12  * attribute. To specify field that should be updated, pass the tab id and
 13  * element id, separated with a colon.
 14  *
 15  * Example 1: (Browse)
 16  *
 17  * <pre>
 18  * {
 19  * 	type : 'button',
 20  * 	id : 'browse',
 21  * 	filebrowser : 'tabId:elementId',
 22  * 	label : editor.lang.common.browseServer
 23  * }
 24  * </pre>
 25  *
 26  * If you set the 'filebrowser' attribute on any element other than
 27  * 'fileButton', the 'Browse' action will be triggered.
 28  *
 29  * Example 2: (Quick Upload)
 30  *
 31  * <pre>
 32  * {
 33  * 	type : 'fileButton',
 34  * 	id : 'uploadButton',
 35  * 	filebrowser : 'tabId:elementId',
 36  * 	label : editor.lang.common.uploadSubmit,
 37  * 	'for' : [ 'upload', 'upload' ]
 38  * }
 39  * </pre>
 40  *
 41  * If you set the 'filebrowser' attribute on a fileButton element, the
 42  * 'QuickUpload' action will be executed.
 43  *
 44  * Filebrowser plugin also supports more advanced configuration (through
 45  * javascript object).
 46  *
 47  * The following settings are supported:
 48  *
 49  * <pre>
 50  *  [action] - Browse or QuickUpload
 51  *  [target] - field to update, tabId:elementId
 52  *  [params] - additional arguments to be passed to the server connector (optional)
 53  *  [onSelect] - function to execute when file is selected/uploaded (optional)
 54  *  [url] - the URL to be called (optional)
 55  * </pre>
 56  *
 57  * Example 3: (Quick Upload)
 58  *
 59  * <pre>
 60  * {
 61  * 	type : 'fileButton',
 62  * 	label : editor.lang.common.uploadSubmit,
 63  * 	id : 'buttonId',
 64  * 	filebrowser :
 65  * 	{
 66  * 		action : 'QuickUpload', //required
 67  * 		target : 'tab1:elementId', //required
 68  * 		params : //optional
 69  * 		{
 70  * 			type : 'Files',
 71  * 			currentFolder : '/folder/'
 72  * 		},
 73  * 		onSelect : function( fileUrl, errorMessage ) //optional
 74  * 		{
 75  * 			// Do not call the built-in selectFuntion
 76  * 			// return false;
 77  * 		}
 78  * 	},
 79  * 	'for' : [ 'tab1', 'myFile' ]
 80  * }
 81  * </pre>
 82  *
 83  * Suppose we have a file element with id 'myFile', text field with id
 84  * 'elementId' and a fileButton. If filebowser.url is not specified explicitly,
 85  * form action will be set to 'filebrowser[DialogName]UploadUrl' or, if not
 86  * specified, to 'filebrowserUploadUrl'. Additional parameters from 'params'
 87  * object will be added to the query string. It is possible to create your own
 88  * uploadHandler and cancel the built-in updateTargetElement command.
 89  *
 90  * Example 4: (Browse)
 91  *
 92  * <pre>
 93  * {
 94  * 	type : 'button',
 95  * 	id : 'buttonId',
 96  * 	label : editor.lang.common.browseServer,
 97  * 	filebrowser :
 98  * 	{
 99  * 		action : 'Browse',
100  * 		url : '/ckfinder/ckfinder.html&type=Images',
101  * 		target : 'tab1:elementId'
102  * 	}
103  * }
104  * </pre>
105  *
106  * In this example, after pressing a button, file browser will be opened in a
107  * popup. If we don't specify filebrowser.url attribute,
108  * 'filebrowser[DialogName]BrowseUrl' or 'filebrowserBrowseUrl' will be used.
109  * After selecting a file in a file browser, an element with id 'elementId' will
110  * be updated. Just like in the third example, a custom 'onSelect' function may be
111  * defined.
112  */
113 ( function()
114 {
115 	/**
116 	 * Adds (additional) arguments to given url.
117 	 *
118 	 * @param {String}
119 	 *            url The url.
120 	 * @param {Object}
121 	 *            params Additional parameters.
122 	 */
123 	function addQueryString( url, params )
124 	{
125 		var queryString = [];
126 
127 		if ( !params )
128 			return url;
129 		else
130 		{
131 			for ( var i in params )
132 				queryString.push( i + "=" + encodeURIComponent( params[ i ] ) );
133 		}
134 
135 		return url + ( ( url.indexOf( "?" ) != -1 ) ? "&" : "?" ) + queryString.join( "&" );
136 	}
137 
138 	/**
139 	 * Make a string's first character uppercase.
140 	 *
141 	 * @param {String}
142 	 *            str String.
143 	 */
144 	function ucFirst( str )
145 	{
146 		str += '';
147 		var f = str.charAt( 0 ).toUpperCase();
148 		return f + str.substr( 1 );
149 	}
150 
151 	/**
152 	 * The onlick function assigned to the 'Browse Server' button. Opens the
153 	 * file browser and updates target field when file is selected.
154 	 *
155 	 * @param {CKEDITOR.event}
156 	 *            evt The event object.
157 	 */
158 	function browseServer( evt )
159 	{
160 		var dialog = this.getDialog();
161 		var editor = dialog.getParentEditor();
162 
163 		editor._.filebrowserSe = this;
164 
165 		var width = editor.config[ 'filebrowser' + ucFirst( dialog.getName() ) + 'WindowWidth' ]
166 				|| editor.config.filebrowserWindowWidth || '80%';
167 		var height = editor.config[ 'filebrowser' + ucFirst( dialog.getName() ) + 'WindowHeight' ]
168 				|| editor.config.filebrowserWindowHeight || '70%';
169 
170 		var params = this.filebrowser.params || {};
171 		params.CKEditor = editor.name;
172 		params.CKEditorFuncNum = editor._.filebrowserFn;
173 		if ( !params.langCode )
174 			params.langCode = editor.langCode;
175 
176 		var url = addQueryString( this.filebrowser.url, params );
177 		editor.popup( url, width, height, editor.config.fileBrowserWindowFeatures );
178 	}
179 
180 	/**
181 	 * The onlick function assigned to the 'Upload' button. Makes the final
182 	 * decision whether form is really submitted and updates target field when
183 	 * file is uploaded.
184 	 *
185 	 * @param {CKEDITOR.event}
186 	 *            evt The event object.
187 	 */
188 	function uploadFile( evt )
189 	{
190 		var dialog = this.getDialog();
191 		var editor = dialog.getParentEditor();
192 
193 		editor._.filebrowserSe = this;
194 
195 		// If user didn't select the file, stop the upload.
196 		if ( !dialog.getContentElement( this[ 'for' ][ 0 ], this[ 'for' ][ 1 ] ).getInputElement().$.value )
197 			return false;
198 
199 		if ( !dialog.getContentElement( this[ 'for' ][ 0 ], this[ 'for' ][ 1 ] ).getAction() )
200 			return false;
201 
202 		return true;
203 	}
204 
205 	/**
206 	 * Setups the file element.
207 	 *
208 	 * @param {CKEDITOR.ui.dialog.file}
209 	 *            fileInput The file element used during file upload.
210 	 * @param {Object}
211 	 *            filebrowser Object containing filebrowser settings assigned to
212 	 *            the fileButton associated with this file element.
213 	 */
214 	function setupFileElement( editor, fileInput, filebrowser )
215 	{
216 		var params = filebrowser.params || {};
217 		params.CKEditor = editor.name;
218 		params.CKEditorFuncNum = editor._.filebrowserFn;
219 		if ( !params.langCode )
220 			params.langCode = editor.langCode;
221 
222 		fileInput.action = addQueryString( filebrowser.url, params );
223 		fileInput.filebrowser = filebrowser;
224 	}
225 
226 	/**
227 	 * Traverse through the content definition and attach filebrowser to
228 	 * elements with 'filebrowser' attribute.
229 	 *
230 	 * @param String
231 	 *            dialogName Dialog name.
232 	 * @param {CKEDITOR.dialog.dialogDefinitionObject}
233 	 *            definition Dialog definition.
234 	 * @param {Array}
235 	 *            elements Array of {@link CKEDITOR.dialog.contentDefinition}
236 	 *            objects.
237 	 */
238 	function attachFileBrowser( editor, dialogName, definition, elements )
239 	{
240 		var element, fileInput;
241 
242 		for ( var i in elements )
243 		{
244 			element = elements[ i ];
245 
246 			if ( element.type == 'hbox' || element.type == 'vbox' )
247 				attachFileBrowser( editor, dialogName, definition, element.children );
248 
249 			if ( !element.filebrowser )
250 				continue;
251 
252 			if ( typeof element.filebrowser == 'string' )
253 			{
254 				var fb =
255 				{
256 					action : ( element.type == 'fileButton' ) ? 'QuickUpload' : 'Browse',
257 					target : element.filebrowser
258 				};
259 				element.filebrowser = fb;
260 			}
261 
262 			if ( element.filebrowser.action == 'Browse' )
263 			{
264 				var url = element.filebrowser.url;
265 				if ( url === undefined )
266 				{
267 					url = editor.config[ 'filebrowser' + ucFirst( dialogName ) + 'BrowseUrl' ];
268 					if ( url === undefined )
269 						url = editor.config.filebrowserBrowseUrl;
270 				}
271 
272 				if ( url )
273 				{
274 					element.onClick = browseServer;
275 					element.filebrowser.url = url;
276 					element.hidden = false;
277 				}
278 			}
279 			else if ( element.filebrowser.action == 'QuickUpload' && element[ 'for' ] )
280 			{
281 				var url = element.filebrowser.url;
282 				if ( url === undefined )
283 				{
284 					url = editor.config[ 'filebrowser' + ucFirst( dialogName ) + 'UploadUrl' ];
285 					if ( url === undefined )
286 						url = editor.config.filebrowserUploadUrl;
287 				}
288 
289 				if ( url )
290 				{
291 					var onClick = element.onClick;
292 					element.onClick = function( evt )
293 					{
294 						// "element" here means the definition object, so we need to find the correct
295 						// button to scope the event call
296 						var sender = evt.sender;
297 						if ( onClick && onClick.call( sender, evt ) === false )
298 							return false;
299 
300 						return uploadFile.call( sender, evt );
301 					};
302 
303 					element.filebrowser.url = url;
304 					element.hidden = false;
305 					setupFileElement( editor, definition.getContents( element[ 'for' ][ 0 ] ).get( element[ 'for' ][ 1 ] ), element.filebrowser );
306 				}
307 			}
308 		}
309 	}
310 
311 	/**
312 	 * Updates the target element with the url of uploaded/selected file.
313 	 *
314 	 * @param {String}
315 	 *            url The url of a file.
316 	 */
317 	function updateTargetElement( url, sourceElement )
318 	{
319 		var dialog = sourceElement.getDialog();
320 		var targetElement = sourceElement.filebrowser.target || null;
321 		url = url.replace( /#/g, '%23' );
322 
323 		// If there is a reference to targetElement, update it.
324 		if ( targetElement )
325 		{
326 			var target = targetElement.split( ':' );
327 			var element = dialog.getContentElement( target[ 0 ], target[ 1 ] );
328 			if ( element )
329 			{
330 				element.setValue( url );
331 				dialog.selectPage( target[ 0 ] );
332 			}
333 		}
334 	}
335 
336 	/**
337 	 * Returns true if filebrowser is configured in one of the elements.
338 	 *
339 	 * @param {CKEDITOR.dialog.dialogDefinitionObject}
340 	 *            definition Dialog definition.
341 	 * @param String
342 	 *            tabId The tab id where element(s) can be found.
343 	 * @param String
344 	 *            elementId The element id (or ids, separated with a semicolon) to check.
345 	 */
346 	function isConfigured( definition, tabId, elementId )
347 	{
348 		if ( elementId.indexOf( ";" ) !== -1 )
349 		{
350 			var ids = elementId.split( ";" );
351 			for ( var i = 0 ; i < ids.length ; i++ )
352 			{
353 				if ( isConfigured( definition, tabId, ids[i] ) )
354 					return true;
355 			}
356 			return false;
357 		}
358 
359 		var elementFileBrowser = definition.getContents( tabId ).get( elementId ).filebrowser;
360 		return ( elementFileBrowser && elementFileBrowser.url );
361 	}
362 
363 	function setUrl( fileUrl, data )
364 	{
365 		var dialog = this._.filebrowserSe.getDialog(),
366 			targetInput = this._.filebrowserSe[ 'for' ],
367 			onSelect = this._.filebrowserSe.filebrowser.onSelect;
368 
369 		if ( targetInput )
370 			dialog.getContentElement( targetInput[ 0 ], targetInput[ 1 ] ).reset();
371 
372 		if ( typeof data == 'function' && data.call( this._.filebrowserSe ) === false )
373 			return;
374 
375 		if ( onSelect && onSelect.call( this._.filebrowserSe, fileUrl, data ) === false )
376 			return;
377 
378 		// The "data" argument may be used to pass the error message to the editor.
379 		if ( typeof data == 'string' && data )
380 			alert( data );
381 
382 		if ( fileUrl )
383 			updateTargetElement( fileUrl, this._.filebrowserSe );
384 	}
385 
386 	CKEDITOR.plugins.add( 'filebrowser',
387 	{
388 		init : function( editor, pluginPath )
389 		{
390 			editor._.filebrowserFn = CKEDITOR.tools.addFunction( setUrl, editor );
391 		}
392 	} );
393 
394 	CKEDITOR.on( 'dialogDefinition', function( evt )
395 	{
396 		var definition = evt.data.definition,
397 			element;
398 		// Associate filebrowser to elements with 'filebrowser' attribute.
399 		for ( var i in definition.contents )
400 		{
401 			if ( ( element = definition.contents[ i ] ) )
402 			{
403 				attachFileBrowser( evt.editor, evt.data.name, definition, element.elements );
404 				if ( element.hidden && element.filebrowser )
405 				{
406 					element.hidden = !isConfigured( definition, element[ 'id' ], element.filebrowser );
407 				}
408 			}
409 		}
410 	} );
411 
412 } )();
413 
414 /**
415  * The location of an external file browser, that should be launched when "Browse Server" button is pressed.
416  * If configured, the "Browse Server" button will appear in Link, Image and Flash dialogs.
417  * @see The <a href="http://docs.cksource.com/CKEditor_3.x/Developers_Guide/File_Browser_(Uploader)">File Browser/Uploader</a> documentation.
418  * @name CKEDITOR.config.filebrowserBrowseUrl
419  * @since 3.0
420  * @type String
421  * @default '' (empty string = disabled)
422  * @example
423  * config.filebrowserBrowseUrl = '/browser/browse.php';
424  */
425 
426 /**
427  * The location of a script that handles file uploads.
428  * If set, the "Upload" tab will appear in "Link", "Image" and "Flash" dialogs.
429  * @name CKEDITOR.config.filebrowserUploadUrl
430  * @see The <a href="http://docs.cksource.com/CKEditor_3.x/Developers_Guide/File_Browser_(Uploader)">File Browser/Uploader</a> documentation.
431  * @since 3.0
432  * @type String
433  * @default '' (empty string = disabled)
434  * @example
435  * config.filebrowserUploadUrl = '/uploader/upload.php';
436  */
437 
438 /**
439  * The location of an external file browser, that should be launched when "Browse Server" button is pressed in the Image dialog.
440  * If not set, CKEditor will use {@link CKEDITOR.config.filebrowserBrowseUrl}.
441  * @name CKEDITOR.config.filebrowserImageBrowseUrl
442  * @since 3.0
443  * @type String
444  * @default '' (empty string = disabled)
445  * @example
446  * config.filebrowserImageBrowseUrl = '/browser/browse.php?type=Images';
447  */
448 
449 /**
450  * The location of an external file browser, that should be launched when "Browse Server" button is pressed in the Flash dialog.
451  * If not set, CKEditor will use {@link CKEDITOR.config.filebrowserBrowseUrl}.
452  * @name CKEDITOR.config.filebrowserFlashBrowseUrl
453  * @since 3.0
454  * @type String
455  * @default '' (empty string = disabled)
456  * @example
457  * config.filebrowserFlashBrowseUrl = '/browser/browse.php?type=Flash';
458  */
459 
460 /**
461  * The location of a script that handles file uploads in the Image dialog.
462  * If not set, CKEditor will use {@link CKEDITOR.config.filebrowserUploadUrl}.
463  * @name CKEDITOR.config.filebrowserImageUploadUrl
464  * @since 3.0
465  * @type String
466  * @default '' (empty string = disabled)
467  * @example
468  * config.filebrowserImageUploadUrl = '/uploader/upload.php?type=Images';
469  */
470 
471 /**
472  * The location of a script that handles file uploads in the Flash dialog.
473  * If not set, CKEditor will use {@link CKEDITOR.config.filebrowserUploadUrl}.
474  * @name CKEDITOR.config.filebrowserFlashUploadUrl
475  * @since 3.0
476  * @type String
477  * @default '' (empty string = disabled)
478  * @example
479  * config.filebrowserFlashUploadUrl = '/uploader/upload.php?type=Flash';
480  */
481 
482 /**
483  * The location of an external file browser, that should be launched when "Browse Server" button is pressed in the Link tab of Image dialog.
484  * If not set, CKEditor will use {@link CKEDITOR.config.filebrowserBrowseUrl}.
485  * @name CKEDITOR.config.filebrowserImageBrowseLinkUrl
486  * @since 3.2
487  * @type String
488  * @default '' (empty string = disabled)
489  * @example
490  * config.filebrowserImageBrowseLinkUrl = '/browser/browse.php';
491  */
492 
493 /**
494  * The "features" to use in the file browser popup window.
495  * @name CKEDITOR.config.filebrowserWindowFeatures
496  * @since 3.4.1
497  * @type String
498  * @default 'location=no,menubar=no,toolbar=no,dependent=yes,minimizable=no,modal=yes,alwaysRaised=yes,resizable=yes,scrollbars=yes'
499  * @example
500  * config.filebrowserWindowFeatures = 'resizable=yes,scrollbars=no';
501  */
502 