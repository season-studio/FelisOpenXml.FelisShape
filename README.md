# FelisShape

FelisShape is a .NET library for manipulating the presentations which conform to the Office Open XML File Formats specification. It is a platform-independent library base on the Open XML SDK. 

## Installation
```
dotnet add package FelisShape
```

## Namespace

The basic namespace is ```FelisOpenXml.FelisShape```.
Some class may be contained in sub-namespace such as ```FelisOpenXml.FelisShape.Draw``` and so on.

## Basic Usage

### Presentation

Load a presentation from a stream:
``` C#
var pres = new FelisPresentaion(soureceStream);
```

Create a empty presentation taken an other one as the template:
``` C#
var template = new FelisPresentaion(templateStream);
var target = FelisPresentation.From(template);
```

Save a presentation:
``` C#
// pres is an instance of FelisPresentation
pres.Save(targetStream);	// save to a stream
pres.Save("filePath.pptx");	// save to a file
```

### Slide

Enumerating the slides in a presentation
``` C#
// pres is an instance of FelisPresentation
foreach (var slide in pres.Slides)
{
	// TODO: ...
}
```

Insert a duplicate of a slide into a presentation
``` C#
pres.InsertSlide(sourceSlide, indexForInsertingAt);
```

Remove a slide in a presentation
``` C#
slide.Remove();
// or
pres.RemoveSlide(slide);
```

Edit a customer data in a slide
``` C#
slide.WorkWithCustomerData("nameOfTheData", (XmlDocument dataDoc) =>
{
	// TODO: ...
	return true; // Return true means there is some changings should be commited to the slide. Otherwise return false.
}, true);
// The last argument of WorkWithCustomerData means if a new customer data should be created when there is no existing one.
// The default value is false.
```

Remove a customer data in a slide
``` C#
slide.RemoveCustomerData("nameOfTheData");
```

Sumbit any changings in the slide
``` C#
slide.Submit();
// Notice: Without invoking this method, the changings in the slide may be lost after close the presentation.
```

### Shapes

Enumerating the shapes in a slide
``` C#
foreach (var shape in slide.Shapes)
{
	// TODO: ...
}
```

Get a shape by a given ID in a slide or in a shapes group
``` C#
slideOrGroup.GetShapeById(id);
```

Get a shape with a special type a slide or in a shapes group
``` C#
slideOrGroup.GetShapes<T>();
// The T can be one of follows:
//		FelisShape
//		FelisPicture 
//		FelisTable 
//		FelisChart
//		FelisConnectionShape
//		FelisOleObject
//		FelisShapeGroup
```

### Manipulating the data of the shape

The id of the shape
``` C#
Console.WriteLine(shape.Id);
shape.Id = newId;
```

The name of the shape
``` C#
Console.WriteLine(shape.Name);
shape.Name = newName;
```

The rect of the shape
``` C#
var rect = shape.Rect;
shape.Rect = new FelisShapeRect() { x = 0, y = 0, cx = 100, cy = 100 };
```

The rect of the shape, the coordinate is relative to the parent shape
``` C#
var ret = shape.RelativeRect;
shape.RelativeRect = new FelisShapeRect() { x = 0, y = 0, cx = 100, cy = 100 };
```

The text in the shape
``` C#
Console.WriteLine(shape.TextBody?.Text);
shape.TextBody.Text = "Hello world";
```

The fill of the shape
``` C#
var fill = new FelisSolidFillValue();
fill.Color.Value = Color.FromArgb(255, 0, 0);
shape.Fill = fill;
```

### See the [API document](./Doc/Api.xml) for more information

------
Season Studio Copyright(2023)