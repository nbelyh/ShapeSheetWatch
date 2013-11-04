
#include "stdafx.h"
#include "lib/Utils.h"
#include "ShapeSheet.h"

using namespace Visio;

enum SectionType
{
	SectionType_Named,
	SectionType_Indexed,
	SectionType_Normal,
	SectionType_Geometry,
	SectionType_Tab,

	SectionType_Count,
};

struct SSInfo
{
	LPCWSTR block_name;
	short section;
	LPCWSTR name;
	LPCWSTR type;
	LPCWSTR values;
};

#define countof(a)	(sizeof(a)/sizeof(a[0]))

SSInfo ss_info[] = {

	{L"1-D Endpoints",visSectionObject,L"BeginX",L"",L""},
	{L"1-D Endpoints",visSectionObject,L"BeginY",L"",L""},
	{L"1-D Endpoints",visSectionObject,L"EndX",L"",L""},
	{L"1-D Endpoints",visSectionObject,L"EndY",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,L"KeepTextFlat",L"BOOL",L"TRUE;FALSE"},
	{L"3-D Rotation Properties",visSectionObject,L"Perspective",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,L"RotationType",L"",L"0;1;2;3;4;5;6"},
	{L"3-D Rotation Properties",visSectionObject,L"RotationXAngle",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,L"RotationYAngle",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,L"RotationZAngle",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.ButtonFace",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.Description",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.DisplayMode",L"ENUM",L"0=visSmartTagDispModeMouseOver;1=visSmartTagDispModeShapeSelected;2=visSmartTagDispModeAlways"},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.TagName",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.X",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.XJustify",L"ENUM",L"0=visSmartTagXJustifyLeft;1=visSmartTagXJustifyCenter;2=visSmartTagXJustifyRight"},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.Y",L"",L""},
	{L"Action Tags",visSectionSmartTag,L"SmartTags.{r}.YJustify",L"",L"0=visSmartTagYJustifyTop;1=visSmartTagYJustifyMiddle;2=visSmartTagYJustifyBottom"},
	{L"Actions",visSectionAction,L"Actions.{r}.Action",L"",L""},
	{L"Actions",visSectionAction,L"Actions.{r}.BeginGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,L"Actions.{r}.ButtonFace",L"",L""},
	{L"Actions",visSectionAction,L"Actions.{r}.Checked",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,L"Actions.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,L"Actions.{r}.FlyoutChild",L"",L""},
	{L"Actions",visSectionAction,L"Actions.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,L"Actions.{r}.Menu",L"",L""},
	{L"Actions",visSectionAction,L"Actions.{r}.ReadOnly",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,L"Actions.{r}.SortKey",L"",L""},
	{L"Actions",visSectionAction,L"Actions.{r}.TagName",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"GlowColor",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"GlowColorTrans",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"GlowSize",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"ReflectionBlur",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"ReflectionDist",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"ReflectionSize",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"ReflectionTrans",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SketchAmount",L"",L"0;1-25"},
	{L"Additional Effect Properties",visSectionObject,L"SketchEnabled",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SketchFillChange",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SketchLineChange",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SketchLineWeight",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SketchSeed",L"",L""},
	{L"Additional Effect Properties",visSectionObject,L"SoftEdgesSize",L"",L""},
	{L"Alignment",visSectionObject,L"AlignBottom",L"",L""},
	{L"Alignment",visSectionObject,L"AlignCenter",L"",L""},
	{L"Alignment",visSectionObject,L"AlignLeft",L"",L""},
	{L"Alignment",visSectionObject,L"AlignMiddle",L"",L""},
	{L"Alignment",visSectionObject,L"AlignRight",L"",L""},
	{L"Alignment",visSectionObject,L"AlignTop",L"",L""},
	{L"Annotation",visSectionAnnotation,L"Annotation.Comment[{j}]",L"",L""},
	{L"Annotation",visSectionAnnotation,L"Annotation.Date[{j}]",L"",L""},
	{L"Annotation",visSectionAnnotation,L"Annotation.LangID[{j}]",L"",L""},
	{L"Annotation",visSectionAnnotation,L"Annotation.X[{j}]",L"",L""},
	{L"Annotation",visSectionAnnotation,L"Annotation.Y[{j}]",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelBottomHeight",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelBottomType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",visSectionObject,L"BevelBottomWidth",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelContourColor",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelContourSize",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelDepthColor",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelDepthSize",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelLightingAngle",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelLightingType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12;13;14;15"},
	{L"Bevel Properties",visSectionObject,L"BevelMaterialType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11"},
	{L"Bevel Properties",visSectionObject,L"BevelTopHeight",L"",L""},
	{L"Bevel Properties",visSectionObject,L"BevelTopType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",visSectionObject,L"BevelTopWidth",L"",L""},
	{L"Change Shape Behavior",visSectionObject,L"ReplaceCopyCells",L"",L""},
	{L"Change Shape Behavior",visSectionObject,L"ReplaceLockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",visSectionObject,L"ReplaceLockShapeData",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",visSectionObject,L"ReplaceLockText",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,L"Char.AsianFont[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.Case[{i}]",L"ENUM",L"0=visCaseNormal;1=visCaseAllCaps;2=visCaseInitialCaps"},
	{L"Character",visSectionCharacter,L"Char.Color[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Character",visSectionCharacter,L"Char.ComplexScriptFont[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.ComplexScriptSize[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.DblUnderline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,L"Char.DoubleStrikethrough[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.Font[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.FontScale[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.LangID[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.Letterspace[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.Overline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,L"Char.Pos[{i}]",L"ENUM",L"0=visPosNormal;1=visPosSuper;2=visPosSub"},
	{L"Character",visSectionCharacter,L"Char.Size[{i}]",L"",L""},
	{L"Character",visSectionCharacter,L"Char.Strikethru[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,L"Char.Style[{i}]",L"",L"Bold=visBold;Italic=visItalic;Underline=visUnderLine;Small caps=visSmallCaps"},
	{L"Connection Points",visSectionConnectionPts,L"Connections.D[{i}]",L"",L""},
	{L"Connection Points",visSectionConnectionPts,L"Connections.DirX[{i}]",L"",L""},
	{L"Connection Points",visSectionConnectionPts,L"Connections.Type[{i}]",L"ENUM",L"0=visCnnctTypeInward;1=visCnnctTypeOutward;2=visCnnctTypeInwardOutward"},
	{L"Connection Points",visSectionConnectionPts,L"Connections.X{i}",L"",L""},
	{L"Connection Points",visSectionConnectionPts,L"Connections.Y{i}",L"",L""},
	{L"Connection Points",visSectionConnectionPts,L"Connections.DirY[{i}]",L"",L""},
	{L"Controls",visSectionControls,L"Controls.{r}.CanGlue",L"BOOL",L"TRUE;FALSE"},
	{L"Controls",visSectionControls,L"Controls.{r}.Tip",L"",L""},
	{L"Controls",visSectionControls,L"Controls.{r}.X",L"",L""},
	{L"Controls",visSectionControls,L"Controls.{r}.XCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",visSectionControls,L"Controls.{r}.XDyn",L"",L""},
	{L"Controls",visSectionControls,L"Controls.{r}.Y",L"",L""},
	{L"Controls",visSectionControls,L"Controls.{r}.YCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",visSectionControls,L"Controls.{r}.YDyn",L"",L""},
	{L"Document Properties",visSectionObject,L"AddMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,L"DocLangID",L"",L""},
	{L"Document Properties",visSectionObject,L"DocLockDuplicatePage",L"",L""},
	{L"Document Properties",visSectionObject,L"DocLocReplace",L"",L""},
	{L"Document Properties",visSectionObject,L"LockPreview",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,L"NoCoauth",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,L"OutputFormat",L"ENUM",L"0;1;2"},
	{L"Document Properties",visSectionObject,L"PreviewQuality",L"ENUM",L"0=visDocPreviewQualityDraft;1=visDocPreviewQualityDetailed"},
	{L"Document Properties",visSectionObject,L"PreviewScope",L"ENUM",L"0=visDocPreviewScope1stPage;1=visDocPreviewScopeNone;2=visDocPreviewScopeAllPages"},
	{L"Document Properties",visSectionObject,L"ViewMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Events",visSectionObject,L"EventDblClick",L"",L""},
	{L"Events",visSectionObject,L"EventDrop",L"",L""},
	{L"Events",visSectionObject,L"EventMultiDrop",L"",L""},
	{L"Events",visSectionObject,L"EventXFMod",L"",L""},
	{L"Events",visSectionObject,L"TheText",L"",L""},
	{L"Fill Format",visSectionObject,L"FillBkgnd",L"",L""},
	{L"Fill Format",visSectionObject,L"FillBkgndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",visSectionObject,L"FillForegnd",L"",L""},
	{L"Fill Format",visSectionObject,L"FillForegndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",visSectionObject,L"FillPattern",L"ENUM",L"0;1;2 - 40"},
	{L"Fill Format",visSectionObject,L"ShapeShdwBlur",L"",L""},
	{L"Fill Format",visSectionObject,L"ShapeShdwObliqueAngle",L"",L""},
	{L"Fill Format",visSectionObject,L"ShapeShdwOffsetX",L"",L""},
	{L"Fill Format",visSectionObject,L"ShapeShdwOffsetY",L"",L""},
	{L"Fill Format",visSectionObject,L"ShapeShdwScaleFactor",L"",L""},
	{L"Fill Format",visSectionObject,L"ShapeShdwShow",L"ENUM",L"0;1;2"},
	{L"Fill Format",visSectionObject,L"ShapeShdwType",L"ENUM",L"0=visFSTPageDefault;1=visFSTSimple;2=visFSTOblique"},
	{L"Fill Format",visSectionObject,L"ShdwForegnd",L"",L""},
	{L"Fill Format",visSectionObject,L"ShdwForegndTrans",L"",L"0 - 100"},
	{L"Fill Format",visSectionObject,L"ShdwPattern",L"",L"0;1;2 - 40"},
	{L"Foreign Image Info",visSectionObject,L"ClippingPath",L"",L""},
	{L"Foreign Image Info",visSectionObject,L"ImgHeight",L"",L""},
	{L"Foreign Image Info",visSectionObject,L"ImgOffsetX",L"",L""},
	{L"Foreign Image Info",visSectionObject,L"ImgOffsetY",L"",L""},
	{L"Foreign Image Info",visSectionObject,L"ImgWidth",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.A{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.B{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.C{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.D{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.E{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.NoFill",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.NoLine",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.NoQuickDrag",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.NoShow",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.NoSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.X{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,L"Geometry{i}.Y{j}",L"",L""},
	{L"Glue Info",visSectionObject,L"BegTrigger",L"",L""},
	{L"Glue Info",visSectionObject,L"EndTrigger",L"",L""},
	{L"Glue Info",visSectionObject,L"GlueType",L"ENUM",L"0=visGlueTypeDefault;2=visGlueTypeWalking;4=visGlueTypeNoWalking;8=visGlueTypeNoWalkingTo"},
	{L"Glue Info",visSectionObject,L"WalkPreference",L"ENUM",L"1=visWalkPrefBegNS;2=visWalkPrefEndNS"},
	{L"Gradient Properties",visSectionObject,L"FillGradientAngle",L"",L""},
	{L"Gradient Properties",visSectionObject,L"FillGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",visSectionObject,L"FillGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,L"LineGradientAngle",L"",L""},
	{L"Gradient Properties",visSectionObject,L"LineGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",visSectionObject,L"LineGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,L"RotateGradientWithShape",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,L"UseGroupGradient",L"",L""},
	{L"Group Properties",visSectionObject,L"DisplayMode",L"ENUM",L"0=visGrpDispModeNone;1=visGrpDispModeBack;2=visGrpDispModeFront"},
	{L"Group Properties",visSectionObject,L"DontMoveChildren",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,L"IsDropTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,L"IsSnapTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,L"IsTextEditTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,L"SelectMode",L"",L"0=visGrpSelModeGroupOnly;1=visGrpSelModeGroup1st;2=visGrpSelModeMembers1st"},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.Address",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.Default",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.Description",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.ExtraInfo",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.Frame",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.NewWindow",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.SortKey",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,L"Hyperlink.{r}.SubAddress",L"",L""},
	{L"Image Properties",visSectionObject,L"Blur",L"",L""},
	{L"Image Properties",visSectionObject,L"Brightness",L"",L""},
	{L"Image Properties",visSectionObject,L"Contrast",L"",L""},
	{L"Image Properties",visSectionObject,L"Denoise",L"",L""},
	{L"Image Properties",visSectionObject,L"Gamma",L"",L""},
	{L"Image Properties",visSectionObject,L"Sharpen",L"",L""},
	{L"Image Properties",visSectionObject,L"Transparency",L"",L"0 - 100"},
	{L"Layers",visSectionLayer,L"Layers.Active[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,L"Layers.Color[{i}]",L"",L""},
	{L"Layers",visSectionLayer,L"Layers.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Layers",visSectionLayer,L"Layers.Glue[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,L"Layers.Locked[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,L"Layers.Print[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,L"Layers.Snap[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,L"Layers.Visible[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Line Format",visSectionObject,L"BeginArrow",L"",L"0;1 - 45"},
	{L"Line Format",visSectionObject,L"BeginArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",visSectionObject,L"CompoundType",L"",L"0;1;2;3;4"},
	{L"Line Format",visSectionObject,L"EndArrow",L"",L"0;1 - 45"},
	{L"Line Format",visSectionObject,L"EndArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",visSectionObject,L"LineCap",L"",L"0;1;2"},
	{L"Line Format",visSectionObject,L"LineColor",L"",L""},
	{L"Line Format",visSectionObject,L"LineColorTrans",L"",L"0 - 100"},
	{L"Line Format",visSectionObject,L"LinePattern",L"",L"0;1;2 - 23"},
	{L"Line Format",visSectionObject,L"LineWeight",L"",L""},
	{L"Line Format",visSectionObject,L"Rounding",L"",L""},
	{L"Miscellaneous",visSectionObject,L"Calendar",L"",L""},
	{L"Miscellaneous",visSectionObject,L"Comment",L"",L""},
	{L"Miscellaneous",visSectionObject,L"DropOnPageScale",L"",L""},
	{L"Miscellaneous",visSectionObject,L"DynFeedback",L"",L"0=visDynFBDefault;1=visDynFBUCon3Leg;2=visDynFBUCon5Leg"},
	{L"Miscellaneous",visSectionObject,L"HideText",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"IsDropSource",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"LangID",L"",L""},
	{L"Miscellaneous",visSectionObject,L"LocalizeMerge",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NoAlignBox",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NoCtlHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NoLiveDynamics",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NonPrinting",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NoObjHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"NoProofing",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,L"ObjType",L"",L"0=visLOFlagsVisDecides;1=visLOFlagsPlacable;2=visLOFlagsRoutable;4=visLOFlagsDont;8=visLOFlagsPNRGroup"},
	{L"Miscellaneous",visSectionObject,L"UpdateAlignBox",L"",L""},
	{L"Page Layout",visSectionObject,L"AvenueSizeY",L"",L""},
	{L"Page Layout",visSectionObject,L"AvenueSizeY",L"",L""},
	{L"Page Layout",visSectionObject,L"AvoidPageBreaks",L"",L""},
	{L"Page Layout",visSectionObject,L"BlockSizeX",L"",L""},
	{L"Page Layout",visSectionObject,L"BlockSizeY",L"",L""},
	{L"Page Layout",visSectionObject,L"CtrlAsInput",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,L"DynamicsOff",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,L"EnableGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,L"LineAdjustFrom",L"",L"0=visPLOLineAdjustFromNotRelated;1=visPLOLineAdjustFromAll;2=visPLOLineAdjustFromNone;3=visPLOLineAdjustFromRoutingDefault"},
	{L"Page Layout",visSectionObject,L"LineAdjustTo",L"",L"0=visPLOLineAdjustToDefault;1=visPLOLineAdjustToAll;2=visPLOLineAdjustToNone;3=visPLOLineAdjustToRelated"},
	{L"Page Layout",visSectionObject,L"LineJumpCode",L"",L"0=visPLOJumpNone;1=visPLOJumpHorizontal;2=visPLOJumpVertical;3=visPLOJumpLastRouted;4=visPLOJumpDisplayOrder;5=visPLOJumpReverseDisplayOrder"},
	{L"Page Layout",visSectionObject,L"LineJumpFactorX",L"",L""},
	{L"Page Layout",visSectionObject,L"LineJumpFactorY",L"",L""},
	{L"Page Layout",visSectionObject,L"LineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Page Layout",visSectionObject,L"LineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Page Layout",visSectionObject,L"LineToLineX",L"",L""},
	{L"Page Layout",visSectionObject,L"LineToLineY",L"",L""},
	{L"Page Layout",visSectionObject,L"LineToNodeX",L"",L""},
	{L"Page Layout",visSectionObject,L"LineToNodeY",L"",L""},
	{L"Page Layout",visSectionObject,L"PageLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Page Layout",visSectionObject,L"PageLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Page Layout",visSectionObject,L"PageShapeSplit",L"",L"0=visPLOSplitNone;1=visPLOSplitAllow"},
	{L"Page Layout",visSectionObject,L"PlaceDepth",L"",L"0=visPLOPlaceDepthDefault;1=visPLOPlaceDepthMedium;2=visPLOPlaceDepthDeep;3=visPLOPlaceDepthShallow"},
	{L"Page Layout",visSectionObject,L"PlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Page Layout",visSectionObject,L"PlaceStyle",L"",L""},
	{L"Page Layout",visSectionObject,L"PlowCode",L"",L"0=visPLOPlowNone;1=visPLOPlowAll"},
	{L"Page Layout",visSectionObject,L"ResizePage",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,L"RouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Page Properties",visSectionObject,L"DrawingResizeType",L"",L""},
	{L"Page Properties",visSectionObject,L"DrawingScale",L"",L""},
	{L"Page Properties",visSectionObject,L"DrawingScaleType",L"",L"0=visNoScale;1=visArchitectural;2=visEngineering;3=visScaleCustom;4=visScaleMetric;5=visScaleMechanical"},
	{L"Page Properties",visSectionObject,L"DrawingSizeType",L"",L"0=visPrintSetup;1=visTight;2=visStandard;3=visCustom;4=visLogical;5=visDSMetric;6=visDSEngr;7=visDSArch"},
	{L"Page Properties",visSectionObject,L"InhibitSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",visSectionObject,L"PageHeight",L"",L""},
	{L"Page Properties",visSectionObject,L"PageLockDuplicate",L"",L""},
	{L"Page Properties",visSectionObject,L"PageLockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",visSectionObject,L"PageScale",L"",L""},
	{L"Page Properties",visSectionObject,L"PageWidth",L"",L""},
	{L"Page Properties",visSectionObject,L"ShdwObliqueAngle",L"",L""},
	{L"Page Properties",visSectionObject,L"ShdwOffsetX",L"",L""},
	{L"Page Properties",visSectionObject,L"ShdwOffsetY",L"",L""},
	{L"Page Properties",visSectionObject,L"ShdwScaleFactor",L"",L""},
	{L"Page Properties",visSectionObject,L"ShdwType",L"",L"1=visFSTSimple;2=visFSTOblique;3=visFSTInner"},
	{L"Page Properties",visSectionObject,L"UIVisibility",L"",L"0=visUIVNormal;1=visUIVHidden"},
	{L"Paragraph",visSectionParagraph,L"Para.Bullet[{i}]",L"",L"0;1;2;3;4;5;6;7"},
	{L"Paragraph",visSectionParagraph,L"Para.BulletFont[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.BulletFontSize[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.BulletStr[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.Flags[{i}]",L"",L"0;1"},
	{L"Paragraph",visSectionParagraph,L"Para.HorzAlign[{i}]",L"",L"0=visHorzLeft;1=visHorzCenter;2=visHorzRight;3=visHorzJustify;4=visHorzForce"},
	{L"Paragraph",visSectionParagraph,L"Para.IndFirst[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.IndLeft[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.IndRight[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.SpAfter[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.SpBefore[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.SpLine[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,L"Para.TextPosAfterBullet[{i}]",L"",L""},
	{L"Print Properties",visSectionObject,L"CenterX",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,L"CenterY",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,L"OnPage",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,L"PageBottomMargin",L"",L""},
	{L"Print Properties",visSectionObject,L"PageLeftMargin",L"",L""},
	{L"Print Properties",visSectionObject,L"PageRightMargin",L"",L""},
	{L"Print Properties",visSectionObject,L"PagesX",L"",L""},
	{L"Print Properties",visSectionObject,L"PagesY",L"",L""},
	{L"Print Properties",visSectionObject,L"PageTopMargin",L"",L""},
	{L"Print Properties",visSectionObject,L"PaperKind",L"",L""},
	{L"Print Properties",visSectionObject,L"PrintGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,L"PrintPageOrientation",L"",L"0=visPPOSameAsPrinter;1=visPPOPortrait;2=visPPOLandscape"},
	{L"Print Properties",visSectionObject,L"ScaleX",L"",L""},
	{L"Print Properties",visSectionObject,L"ScaleY",L"",L""},
	{L"PrintProperties",visSectionObject,L"PaperSource",L"",L""},
	{L"Protection",visSectionObject,L"LockAspect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockBegin",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockCalcWH",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockCrop",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockCustProp",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockDelete",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockEnd",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockFromGroupFormat",L"",L""},
	{L"Protection",visSectionObject,L"LockGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockHeight",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockMoveX",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockMoveY",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockRotate",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockSelect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockTextEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockThemeColors",L"",L""},
	{L"Protection",visSectionObject,L"LockThemeConnectors",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockThemeEffects",L"",L""},
	{L"Protection",visSectionObject,L"LockThemeFonts",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockThemeIndex",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockVariation",L"",L""},
	{L"Protection",visSectionObject,L"LockVtxEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,L"LockWidth",L"BOOL",L"TRUE;FALSE"},
	{L"Quick Style",visSectionObject,L"QuickStyleEffectsMatrix",L"",L""},
	{L"Quick Style",visSectionObject,L"QuickStyleFillColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,L"QuickStyleFillMatrix",L"",L""},
	{L"Quick Style",visSectionObject,L"QuickStyleFontColor",L"",L"0;1"},
	{L"Quick Style",visSectionObject,L"QuickStyleFontMatrix",L"",L""},
	{L"Quick Style",visSectionObject,L"QuickStyleLineColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,L"QuickStyleLineMatrix",L"",L""},
	{L"Quick Style",visSectionObject,L"QuickStyleShadowColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,L"QuickStyleType",L"",L"0;1;2;3"},
	{L"Quick Style",visSectionObject,L"QuickStyleVariation",L"",L"0;1;2;4;8"},
	{L"Reviewer",visSectionReviewer,L"Reviewer.Color[{i}]",L"",L""},
	{L"Reviewer",visSectionReviewer,L"Reviewer.Initials[{i}]",L"",L""},
	{L"Reviewer",visSectionReviewer,L"Reviewer.Name[{i}]",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"XGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",visSectionObject,L"XGridOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"XGridSpacing",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"XRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",visSectionObject,L"XRulerOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"YGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",visSectionObject,L"YGridOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"YGridSpacing",L"",L""},
	{L"Ruler & Grid",visSectionObject,L"YRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",visSectionObject,L"YRulerOrigin",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Ask",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Calendar",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Format",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Label",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.LangID",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Prompt",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.SortKey",L"",L""},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Type",L"",L"0=visPropTypeString;1=visPropTypeListFix;2=visPropTypeNumber;3=visPropTypeBool;4=visPropTypeListVar;5=visPropTypeDate;6=visPropTypeDuration;7=visPropTypeCurrency"},
	{L"Shape Data",visSectionProp,L"Prop.{r}.Value",L"",L""},
	{L"Shape Layout",visSectionObject,L"ConFixedCode",L"",L"0=visSLOConFixedRerouteFreely;1=visSLOConFixedRerouteAsNeeded;2=visSLOConFixedRerouteNever;3=visSLOConFixedRerouteOnCrossover;4=visSLOConFixedByAlgFrom;5=visSLOConFixedByAlgTo;6=visSLOConFixedByAlgFromTo"},
	{L"Shape Layout",visSectionObject,L"ConLineJumpCode",L"",L"0=visSLOJumpDefault;1=visSLOJumpNever;2=visSLOJumpAlways;3=visSLOJumpOther;4=visSLOJumpNeither"},
	{L"Shape Layout",visSectionObject,L"ConLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Shape Layout",visSectionObject,L"ConLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Shape Layout",visSectionObject,L"ConLineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Shape Layout",visSectionObject,L"ConLineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Shape Layout",visSectionObject,L"DisplayLevel",L"",L""},
	{L"Shape Layout",visSectionObject,L"Relationships",L"",L""},
	{L"Shape Layout",visSectionObject,L"ShapeFixedCode",L"",L"1=visSLOFixedPlacement;2=visSLOFixedPlow;4=visSLOFixedPermeablePlow;32=visSLOFixedConnPtsIgnore;64=visSLOFixedConnPtsOnly;128=visSLOFixedNoFoldToShape"},
	{L"Shape Layout",visSectionObject,L"ShapePermeablePlace",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,L"ShapePermeableX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,L"ShapePermeableY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,L"ShapePlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Shape Layout",visSectionObject,L"ShapePlaceStyle",L"",L"visLOPlaceBottomToTop;visLOPlaceCircular;visLOPlaceCompactDownLeft;visLOPlaceCompactDownRight;visLOPlaceCompactLeftDown;visLOPlaceCompactLeftUp;visLOPlaceCompactRightDown;visLOPlaceCompactRightUp;visLOPlaceCompactUpLeft;visLOPlaceCompactUpRight;visLOPlaceDefault;visLOPlaceHierarchyBottomToTopCenter;visLOPlaceHierarchyBottomToTopLeft;visLOPlaceHierarchyBottomToTopRight;visLOPlaceHierarchyLeftToRightBottom;visLOPlaceHierarchyLeftToRightMiddle;visLOPlaceHierarchyLeftToRightTop;visLOPlaceHierarchyRightToLeftBottom;visLOPlaceHierarchyRightToLeftMiddle;visLOPlaceHierarchyRightToLeftTop;visLOPlaceHierarchyTopToBottomCenter;visLOPlaceHierarchyTopToBottomLeft;visLOPlaceHierarchyTopToBottomRight;visLOPlaceLeftToRight;visLOPlaceParentDefault;visLOPlaceRadial;visLOPlaceRightToLeft;visLOPlaceTopToBottom"},
	{L"Shape Layout",visSectionObject,L"ShapePlowCode",L"",L"0=visSLOPlowDefault;1=visSLOPlowNever;2=visSLOPlowAlways"},
	{L"Shape Layout",visSectionObject,L"ShapeRouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Shape Layout",visSectionObject,L"ShapeSplit",L"",L"0=visSLOSplitNone;1=visSLOSplitAllow"},
	{L"Shape Layout",visSectionObject,L"ShapeSplittable",L"",L"0=visSLOSplittableNone;1=visSLOSplittableAllow"},
	{L"Shape Transform",visSectionObject,L"Angle",L"",L""},
	{L"Shape Transform",visSectionObject,L"FlipX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",visSectionObject,L"FlipY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",visSectionObject,L"Height",L"",L""},
	{L"Shape Transform",visSectionObject,L"LocPinX",L"",L""},
	{L"Shape Transform",visSectionObject,L"LocPinY",L"",L""},
	{L"Shape Transform",visSectionObject,L"PinX",L"",L""},
	{L"Shape Transform",visSectionObject,L"PinY",L"",L""},
	{L"Shape Transform",visSectionObject,L"ResizeMode",L"",L"0=visXFormResizeDontCare;1=visXFormResizeSpread;2=visXFormResizeScale"},
	{L"Shape Transform",visSectionObject,L"Width",L"",L""},
	{L"Style Properties",visSectionObject,L"EnableFillProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,L"EnableLineProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,L"EnableTextProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,L"HideForApply",L"BOOL",L"TRUE;FALSE"},
	{L"Tabs",visSectionTab,L"Tabs.{i}{j}",L"",L"0=visTabStopLeft;1=visTabStopCenter;2=visTabStopRight;3=visTabStopDecimal;4=visTabStopComma"},
	{L"Tabs",visSectionTab,L"Tabs.{i}{j}",L"",L""},
	{L"Text Block Format",visSectionObject,L"BottomMargin",L"",L""},
	{L"Text Block Format",visSectionObject,L"DefaultTabstop",L"",L""},
	{L"Text Block Format",visSectionObject,L"LeftMargin",L"",L""},
	{L"Text Block Format",visSectionObject,L"RightMargin",L"",L""},
	{L"Text Block Format",visSectionObject,L"TextBkgnd",L"",L""},
	{L"Text Block Format",visSectionObject,L"TextBkgndTrans",L"",L"0 - 100"},
	{L"Text Block Format",visSectionObject,L"TextDirection",L"",L"0=visTxtBlkLeftToRight;1=visTxtBlkTopToBottom"},
	{L"Text Block Format",visSectionObject,L"TopMargin",L"",L""},
	{L"Text Block Format",visSectionObject,L"VerticalAlign",L"",L"0=visVertTop;1=visVertMiddle;2=visVertBottom"},
	{L"Text Fields",visSectionTextField,L"Fields.Calendar[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,L"Fields.Format[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,L"Fields.ObjectKind[{i}]",L"",L"0=visTFOKStandard;1=visTFOKHorizontaInVertical"},
	{L"Text Fields",visSectionTextField,L"Fields.Type[{i}]",L"",L"0;2;5;6;7"},
	{L"Text Fields",visSectionTextField,L"Fields.UICat[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,L"Fields.UICod[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,L"Fields.UIFmt[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,L"Fields.Value[{i}]",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtAngle",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtHeight",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtLocPinX",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtLocPinY",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtPinX",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtPinY",L"",L""},
	{L"Text Transform",visSectionObject,L"TxtWidth",L"",L""},
	{L"Theme Properties",visSectionObject,L"ColorSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"ConnectorSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"EffectSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"EmbellishmentIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"FontSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"ThemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"VariationColorIndex",L"",L""},
	{L"Theme Properties",visSectionObject,L"VariationStyleIndex",L"",L""},
	{L"User-Defined Cells",visSectionUser,L"User.{r}.Prompt",L"",L""},
	{L"User-Defined Cells",visSectionUser,L"User.{r}.Value",L"",L""},
	{L"",visSectionObject,L"DistanceFromGround",L"",L""},
};

void GetVariableIndexedSectionCellNames(IVShapePtr shape, short section_no, const CString& mask, Strings& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	for (short r = 0; r < section->Count; ++r)
	{
		for (size_t i = 0; i < countof(ss_info); ++i)
		{
			if (section_no == ss_info[i].section)
			{
				CString name = ss_info[i].name;
				name.Replace(L"{i}", FormatString(L"%d", 1 + r));

				if (StringIsLike(mask, name))
					result.push_back(name);
			}
		}
	}
}

void GetVariableNamedSectionCellNames(IVShapePtr shape, short section_no, const CString& mask, Strings& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	for (short r = 0; r < section->Count; ++r)
 	{
		IVRowPtr row = section->GetRow(r);
		CString row_name = row->NameU;

		for (size_t i = 0; i < countof(ss_info); ++i)
		{
			if (section_no == ss_info[i].section)
			{
				CString name = ss_info[i].name;
				name.Replace(L"{r}", row_name);

				if (StringIsLike(mask, name))
					result.push_back(name);
			}
		}
	}
}

void GetVariableGeometrySectionCellNames(IVShapePtr shape, const CString& mask, Strings& result)
{
	for (long s = 0; s < shape->GeometryCount; ++s)
	{
		IVSectionPtr section = shape->GetSection(visSectionFirstComponent + s);

		for (short r = 0; r < section->Count; ++r)
		{
			for (size_t i = 0; i < countof(ss_info); ++i)
			{
				if (ss_info[i].section == visSectionFirstComponent)
				{
					CString name = ss_info[i].name;
					name.Replace(L"{i}", FormatString(L"%d", 1 + s));
					name.Replace(L"{j}", FormatString(L"%d", 1 + r));

					if (StringIsLike(mask, name))
						result.push_back(name);
				}
			}
		}
	}
}

void GetCellNames(IVShapePtr shape, const CString& cell_name_mask, Strings& result)
{
	if (!shape)
		return;

	GetVariableNamedSectionCellNames(shape, visSectionAction, cell_name_mask, result);
	GetVariableNamedSectionCellNames(shape, visSectionSmartTag, cell_name_mask, result);
	GetVariableNamedSectionCellNames(shape, visSectionControls, cell_name_mask, result);
	GetVariableNamedSectionCellNames(shape, visSectionHyperlink, cell_name_mask, result);
	GetVariableNamedSectionCellNames(shape, visSectionProp, cell_name_mask, result);
	GetVariableNamedSectionCellNames(shape, visSectionUser, cell_name_mask, result);

	GetVariableIndexedSectionCellNames(shape, visSectionAnnotation, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionCharacter, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionConnectionPts, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionLayer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionParagraph, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionReviewer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionTextField, cell_name_mask, result);

	GetVariableGeometrySectionCellNames(shape, cell_name_mask, result);

}


