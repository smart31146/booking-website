/*
Copyright (c) 2003-2010, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	// Define changes to default configuration here. For example:
	config.language = 'en';
	config.enterMode = CKEDITOR.ENTER_BR;
	config.toolbar = 'MyToolbar';
	config.toolbar = 'basic';
	config.height = "400";
	config.width = "650";
	config.filebrowserBrowseUrl = 'ckeditor/ckfinder/ckfinder.html';
	config.filebrowserImageBrowseUrl = 'ckeditor/ckfinder/ckfinder.html?type=Images';
	config.filebrowserUploadUrl = 'ckeditor/ckfinder/core/connector/asp/connector.asp?command=QuickUpload&type=Files';
	config.filebrowserImageUploadUrl = 'ckeditor/ckfinder/core/connector/asp/connector.asp?command=QuickUpload&type=Images';
    config.toolbar_MyToolbar =
    [
        ['Source','-','PasteFromWord','Undo','Redo','-','Replace','-','SelectAll','RemoveFormat','Table','SpecialChar','CreateDiv'],
        ['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock','BulletedList','-','Subscript','Superscript'],
        '/',
        ['Format','Font','FontSize','TextColor','-','Bold','Italic','Underline','Strike'],
        ['Image','Link','Unlink'],
    ];
	 config.toolbar_basic =
    [
        ['Source','-','PasteFromWord','RemoveFormat','Table','BulletedList','-','TextColor','-','Bold','Italic','Underline','Image','Link','Unlink'],
    ];
};
