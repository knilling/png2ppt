/*
 *Copyright (c) 2016 Christopher Crawford
 *
 *Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 *The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 *THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */

var powerpnt = new ActiveXObject("PowerPoint.Application")
powerpnt.Visible = true;
var ppt = powerpnt.Presentations.Add();
var slide_height = ppt.PageSetup.SlideHeight;
var slide_width = ppt.PageSetup.SlideWidth;
var slides = ppt.Slides;

var fs = new ActiveXObject("Scripting.FileSystemObject");
var sh = new ActiveXObject("WScript.Shell");
var script_name = WScript.ScriptFullName;
sh.CurrentDirectory = fs.GetParentFolderName(script_name);
var photo_folder = fs.GetFolder(sh.CurrentDirectory);

var files = new Enumerator(photo_folder.Files);
var i = 1;
while(!files.atEnd()){
	var file = files.item();
	if(file.Name.match(/\.png$|\.jpg/gi)) {
		//WScript.Echo(file.Path);
		var slide = slides.Add(i,12);
		i = i + 1;
		var my_slide = slide.Shapes;
		var pic = my_slide.AddPicture(file.Path,0,-1,50,100);
		pic.ScaleWidth(.78,-1);
	}
	files.moveNext();
}