
#include "stdafx.h"
#include "import/VISLIB.tlh"
#include "lib/Utils.h"
#include "ShapeSheet.h"

//TODO: tab support {i}{j}

struct SSInfo
{
	LPCWSTR s_name;
	LPCWSTR r_name;
	LPCWSTR c_name;
	short s;
	short r;
	short c;
	LPCWSTR name;
	LPCWSTR type;
	LPCWSTR values;
};

SSInfo ss_info[] = {

	{L"1-D Endpoints",L"",L"BeginX",visSectionObject,visRowXForm1D,vis1DBeginX,L"BeginX",L"",L""},
	{L"1-D Endpoints",L"",L"BeginY",visSectionObject,visRowXForm1D,vis1DBeginY,L"BeginY",L"",L""},
	{L"1-D Endpoints",L"",L"EndX",visSectionObject,visRowXForm1D,vis1DEndX,L"EndX",L"",L""},
	{L"1-D Endpoints",L"",L"EndY",visSectionObject,visRowXForm1D,vis1DEndY,L"EndY",L"",L""},
	{L"3-D Rotation Properties",L"",L"DistanceFromGround",visSectionObject,visRow3DRotationProperties,visDistanceFromGround,L"DistanceFromGround",L"",L""},
	{L"3-D Rotation Properties",L"",L"KeepTextFlat",visSectionObject,visRow3DRotationProperties,visKeepTextFlat,L"KeepTextFlat",L"BOOL",L"TRUE;FALSE"},
	{L"3-D Rotation Properties",L"",L"Perspective",visSectionObject,visRow3DRotationProperties,visPerspective,L"Perspective",L"",L""},
	{L"3-D Rotation Properties",L"",L"RotationType",visSectionObject,visRow3DRotationProperties,visRotationType,L"RotationType",L"",L"0;1;2;3;4;5;6"},
	{L"3-D Rotation Properties",L"",L"RotationXAngle",visSectionObject,visRow3DRotationProperties,visRotationXAngle,L"RotationXAngle",L"",L""},
	{L"3-D Rotation Properties",L"",L"RotationYAngle",visSectionObject,visRow3DRotationProperties,visRotationYAngle,L"RotationYAngle",L"",L""},
	{L"3-D Rotation Properties",L"",L"RotationZAngle",visSectionObject,visRow3DRotationProperties,visRotationZAngle,L"RotationZAngle",L"",L""},
	{L"Action Tags",L"{r}",L"ButtonFace",visSectionSmartTag,visRowSmartTag ,visSmartTagButtonFace,L"SmartTags.{r}.ButtonFace",L"",L""},
	{L"Action Tags",L"{r}",L"Description",visSectionSmartTag,visRowSmartTag ,visSmartTagDescription,L"SmartTags.{r}.Description",L"",L""},
	{L"Action Tags",L"{r}",L"Disabled",visSectionSmartTag,visRowSmartTag ,visSmartTagDisabled,L"SmartTags.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Action Tags",L"{r}",L"DisplayMode",visSectionSmartTag,visRowSmartTag ,visSmartTagDisplayMode,L"SmartTags.{r}.DisplayMode",L"ENUM",L"0=visSmartTagDispModeMouseOver;1=visSmartTagDispModeShapeSelected;2=visSmartTagDispModeAlways"},
	{L"Action Tags",L"{r}",L"TagName",visSectionSmartTag,visRowSmartTag ,visSmartTagName,L"SmartTags.{r}.TagName",L"",L""},
	{L"Action Tags",L"{r}",L"X",visSectionSmartTag,visRowSmartTag ,visSmartTagX,L"SmartTags.{r}.X",L"",L""},
	{L"Action Tags",L"{r}",L"XJustify",visSectionSmartTag,visRowSmartTag ,visSmartTagXJustify,L"SmartTags.{r}.XJustify",L"ENUM",L"0=visSmartTagXJustifyLeft;1=visSmartTagXJustifyCenter;2=visSmartTagXJustifyRight"},
	{L"Action Tags",L"{r}",L"Y",visSectionSmartTag,visRowSmartTag ,visSmartTagY,L"SmartTags.{r}.Y",L"",L""},
	{L"Action Tags",L"{r}",L"YJustify",visSectionSmartTag,visRowSmartTag ,visSmartTagYJustify,L"SmartTags.{r}.YJustify",L"",L"0=visSmartTagYJustifyTop;1=visSmartTagYJustifyMiddle;2=visSmartTagYJustifyBottom"},
	{L"Actions",L"{r}",L"Action",visSectionAction,visRowAction ,visActionAction,L"Actions.{r}.Action",L"",L""},
	{L"Actions",L"{r}",L"BeginGroup",visSectionAction,visRowAction ,visActionBeginGroup,L"Actions.{r}.BeginGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",L"{r}",L"ButtonFace",visSectionAction,visRowAction ,visActionButtonFace,L"Actions.{r}.ButtonFace",L"",L""},
	{L"Actions",L"{r}",L"Checked",visSectionAction,visRowAction ,visActionChecked,L"Actions.{r}.Checked",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",L"{r}",L"Disabled",visSectionAction,visRowAction ,visActionDisabled,L"Actions.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",L"{r}",L"FlyoutChild",visSectionAction,visRowAction ,visActionFlyoutChild,L"Actions.{r}.FlyoutChild",L"",L""},
	{L"Actions",L"{r}",L"Invisible",visSectionAction,visRowAction ,visActionInvisible,L"Actions.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",L"{r}",L"Menu",visSectionAction,visRowAction ,visActionMenu,L"Actions.{r}.Menu",L"",L""},
	{L"Actions",L"{r}",L"ReadOnly",visSectionAction,visRowAction ,visActionReadOnly,L"Actions.{r}.ReadOnly",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",L"{r}",L"SortKey",visSectionAction,visRowAction ,visActionSortKey,L"Actions.{r}.SortKey",L"",L""},
	{L"Actions",L"{r}",L"TagName",visSectionAction,visRowAction ,visActionTagName,L"Actions.{r}.TagName",L"",L""},
	{L"Additional Effect Properties",L"",L"GlowColor",visSectionObject,visRowOtherEffectProperties,visGlowColor,L"GlowColor",L"",L""},
	{L"Additional Effect Properties",L"",L"GlowColorTrans",visSectionObject,visRowOtherEffectProperties,visGlowColorTrans,L"GlowColorTrans",L"",L""},
	{L"Additional Effect Properties",L"",L"GlowSize",visSectionObject,visRowOtherEffectProperties,visGlowSize,L"GlowSize",L"",L""},
	{L"Additional Effect Properties",L"",L"ReflectionBlur",visSectionObject,visRowOtherEffectProperties,visReflectionBlur,L"ReflectionBlur",L"",L""},
	{L"Additional Effect Properties",L"",L"ReflectionDist",visSectionObject,visRowOtherEffectProperties,visReflectionDist,L"ReflectionDist",L"",L""},
	{L"Additional Effect Properties",L"",L"ReflectionSize",visSectionObject,visRowOtherEffectProperties,visReflectionSize,L"ReflectionSize",L"",L""},
	{L"Additional Effect Properties",L"",L"ReflectionTrans",visSectionObject,visRowOtherEffectProperties,visReflectionTrans,L"ReflectionTrans",L"",L""},
	{L"Additional Effect Properties",L"",L"SketchAmount",visSectionObject,visRowOtherEffectProperties,visSketchAmount,L"SketchAmount",L"",L"0;1-25"},
	{L"Additional Effect Properties",L"",L"SketchEnabled",visSectionObject,visRowOtherEffectProperties,visSketchEnabled,L"SketchEnabled",L"",L""},
	{L"Additional Effect Properties",L"",L"SketchFillChange",visSectionObject,visRowOtherEffectProperties,visSketchFillChange,L"SketchFillChange",L"",L""},
	{L"Additional Effect Properties",L"",L"SketchLineChange",visSectionObject,visRowOtherEffectProperties,visSketchLineChange,L"SketchLineChange",L"",L""},
	{L"Additional Effect Properties",L"",L"SketchLineWeight",visSectionObject,visRowOtherEffectProperties,visSketchLineWeight,L"SketchLineWeight",L"",L""},
	{L"Additional Effect Properties",L"",L"SketchSeed",visSectionObject,visRowOtherEffectProperties,visSketchSeed,L"SketchSeed",L"",L""},
	{L"Additional Effect Properties",L"",L"SoftEdgesSize",visSectionObject,visRowOtherEffectProperties,visSoftEdgesSize,L"SoftEdgesSize",L"",L""},
	{L"Alignment",L"",L"AlignBottom",visSectionObject,visRowAlign,visAlignBottom,L"AlignBottom",L"",L""},
	{L"Alignment",L"",L"AlignCenter",visSectionObject,visRowAlign,visAlignCenter,L"AlignCenter",L"",L""},
	{L"Alignment",L"",L"AlignLeft",visSectionObject,visRowAlign,visAlignLeft,L"AlignLeft",L"",L""},
	{L"Alignment",L"",L"AlignMiddle",visSectionObject,visRowAlign,visAlignMiddle,L"AlignMiddle",L"",L""},
	{L"Alignment",L"",L"AlignRight",visSectionObject,visRowAlign,visAlignRight,L"AlignRight",L"",L""},
	{L"Alignment",L"",L"AlignTop",visSectionObject,visRowAlign,visAlignTop,L"AlignTop",L"",L""},
	{L"Annotation",L"{i}",L"Comment",visSectionAnnotation,visRowAnnotation ,visAnnotationComment,L"Annotation.Comment[{i}]",L"",L""},
	{L"Annotation",L"{i}",L"Date",visSectionAnnotation,visRowAnnotation ,visAnnotationDate,L"Annotation.Date[{i}]",L"",L""},
	{L"Annotation",L"{i}",L"LangID",visSectionAnnotation,visRowAnnotation ,visAnnotationLangID,L"Annotation.LangID[{i}]",L"",L""},
	{L"Annotation",L"{i}",L"X",visSectionAnnotation,visRowAnnotation ,visAnnotationX,L"Annotation.X[{i}]",L"",L""},
	{L"Annotation",L"{i}",L"Y",visSectionAnnotation,visRowAnnotation ,visAnnotationY,L"Annotation.Y[{i}]",L"",L""},
	{L"Bevel Properties",L"",L"BevelBottomHeight",visSectionObject,visRowBevelProperties,visBevelBottomHeight,L"BevelBottomHeight",L"",L""},
	{L"Bevel Properties",L"",L"BevelBottomType",visSectionObject,visRowBevelProperties,visBevelBottomType,L"BevelBottomType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",L"",L"BevelBottomWidth",visSectionObject,visRowBevelProperties,visBevelBottomWidth,L"BevelBottomWidth",L"",L""},
	{L"Bevel Properties",L"",L"BevelContourColor",visSectionObject,visRowBevelProperties,visBevelContourColor,L"BevelContourColor",L"",L""},
	{L"Bevel Properties",L"",L"BevelContourSize",visSectionObject,visRowBevelProperties,visBevelContourSize,L"BevelContourSize",L"",L""},
	{L"Bevel Properties",L"",L"BevelDepthColor",visSectionObject,visRowBevelProperties,visBevelDepthColor,L"BevelDepthColor",L"",L""},
	{L"Bevel Properties",L"",L"BevelDepthSize",visSectionObject,visRowBevelProperties,visBevelDepthSize,L"BevelDepthSize",L"",L""},
	{L"Bevel Properties",L"",L"BevelLightingAngle",visSectionObject,visRowBevelProperties,visBevelLightingAngle,L"BevelLightingAngle",L"",L""},
	{L"Bevel Properties",L"",L"BevelLightingType",visSectionObject,visRowBevelProperties,visBevelLightingType,L"BevelLightingType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12;13;14;15"},
	{L"Bevel Properties",L"",L"BevelMaterialType",visSectionObject,visRowBevelProperties,visBevelMaterialType,L"BevelMaterialType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11"},
	{L"Bevel Properties",L"",L"BevelTopHeight",visSectionObject,visRowBevelProperties,visBevelTopHeight,L"BevelTopHeight",L"",L""},
	{L"Bevel Properties",L"",L"BevelTopType",visSectionObject,visRowBevelProperties,visBevelTopType,L"BevelTopType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",L"",L"BevelTopWidth",visSectionObject,visRowBevelProperties,visBevelTopWidth,L"BevelTopWidth",L"",L""},
	{L"Change Shape Behavior",L"",L"ReplaceCopyCells",visSectionObject,visRowReplaceBehaviors,visReplaceCopyCells,L"ReplaceCopyCells",L"",L""},
	{L"Change Shape Behavior",L"",L"ReplaceLockFormat",visSectionObject,visRowReplaceBehaviors,visReplaceLockFormat,L"ReplaceLockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",L"",L"ReplaceLockShapeData",visSectionObject,visRowReplaceBehaviors,visReplaceLockShapeData,L"ReplaceLockShapeData",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",L"",L"ReplaceLockText",visSectionObject,visRowReplaceBehaviors,visReplaceLockText,L"ReplaceLockText",L"BOOL",L"TRUE;FALSE"},
	{L"Character",L"{i}",L"AsianFont",visSectionCharacter,visRowCharacter ,visCharacterAsianFont,L"Char.AsianFont[{i}]",L"",L""},
	{L"Character",L"{i}",L"Case",visSectionCharacter,visRowCharacter ,visCharacterCase,L"Char.Case[{i}]",L"ENUM",L"0=visCaseNormal;1=visCaseAllCaps;2=visCaseInitialCaps"},
	{L"Character",L"{i}",L"Color",visSectionCharacter,visRowCharacter ,visCharacterColor,L"Char.Color[{i}]",L"",L""},
	{L"Character",L"{i}",L"ColorTrans",visSectionCharacter,visRowCharacter ,visCharacterColorTrans,L"Char.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Character",L"{i}",L"ComplexScriptFont",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptFont,L"Char.ComplexScriptFont[{i}]",L"",L""},
	{L"Character",L"{i}",L"ComplexScriptSize",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptSize,L"Char.ComplexScriptSize[{i}]",L"",L""},
	{L"Character",L"{i}",L"DblUnderline",visSectionCharacter,visRowCharacter ,visCharacterDblUnderline,L"Char.DblUnderline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",L"{i}",L"DoubleStrikethrough",visSectionCharacter,visRowCharacter ,visCharacterDoubleStrikethrough,L"Char.DoubleStrikethrough[{i}]",L"",L""},
	{L"Character",L"{i}",L"Font",visSectionCharacter,visRowCharacter ,visCharacterFont,L"Char.Font[{i}]",L"",L""},
	{L"Character",L"{i}",L"FontScale",visSectionCharacter,visRowCharacter ,visCharacterFontScale,L"Char.FontScale[{i}]",L"",L""},
	{L"Character",L"{i}",L"LangID",visSectionCharacter,visRowCharacter ,visCharacterLangID,L"Char.LangID[{i}]",L"",L""},
	{L"Character",L"{i}",L"Letterspace",visSectionCharacter,visRowCharacter ,visCharacterLetterspace,L"Char.Letterspace[{i}]",L"",L""},
	{L"Character",L"{i}",L"Overline",visSectionCharacter,visRowCharacter ,visCharacterOverline,L"Char.Overline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",L"{i}",L"Pos",visSectionCharacter,visRowCharacter ,visCharacterPos,L"Char.Pos[{i}]",L"ENUM",L"0=visPosNormal;1=visPosSuper;2=visPosSub"},
	{L"Character",L"{i}",L"Size",visSectionCharacter,visRowCharacter ,visCharacterSize,L"Char.Size[{i}]",L"",L""},
	{L"Character",L"{i}",L"Strikethru",visSectionCharacter,visRowCharacter ,visCharacterStrikethru,L"Char.Strikethru[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",L"{i}",L"Style",visSectionCharacter,visRowCharacter ,visCharacterStyle,L"Char.Style[{i}]",L"",L"Bold=visBold;Italic=visItalic;Underline=visUnderLine;Small caps=visSmallCaps"},
	{L"Connection Points",L"{i}",L"D",visSectionConnectionPts,visRowConnectionPts ,visCnnctD,L"Connections.D[{i}]",L"",L""},
	{L"Connection Points",L"{i}",L"DirX",visSectionConnectionPts,visRowConnectionPts ,visCnnctA,L"Connections.DirX[{i}]",L"",L""},
	{L"Connection Points",L"{i}",L"Type",visSectionConnectionPts,visRowConnectionPts ,visCnnctC,L"Connections.Type[{i}]",L"ENUM",L"0=visCnnctTypeInward;1=visCnnctTypeOutward;2=visCnnctTypeInwardOutward"},
	{L"Connection Points",L"{i}",L"X",visSectionConnectionPts,visRowConnectionPts ,visX,L"Connections.X{i}",L"",L""},
	{L"Connection Points",L"{i}",L"Y",visSectionConnectionPts,visRowConnectionPts ,visY,L"Connections.Y{i}",L"",L""},
	{L"Connection Points",L"{i}",L"DirY",visSectionConnectionPts,visRowConnectionPts ,visCnnctB,L"Connections.DirY[{i}]",L"",L""},
	{L"Controls",L"{r}",L"CanGlue",visSectionControls,visRowControl ,visCtlGlue,L"Controls.{r}.CanGlue",L"BOOL",L"TRUE;FALSE"},
	{L"Controls",L"{r}",L"Tip",visSectionControls,visRowControl ,visCtlTip,L"Controls.{r}.Tip",L"",L""},
	{L"Controls",L"{r}",L"X",visSectionControls,visRowControl ,visCtlX,L"Controls.{r}.X",L"",L""},
	{L"Controls",L"{r}",L"XCon",visSectionControls,visRowControl ,visCtlXCon,L"Controls.{r}.XCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",L"{r}",L"XDyn",visSectionControls,visRowControl ,visCtlXDyn,L"Controls.{r}.XDyn",L"",L""},
	{L"Controls",L"{r}",L"Y",visSectionControls,visRowControl ,visCtlY,L"Controls.{r}.Y",L"",L""},
	{L"Controls",L"{r}",L"YCon",visSectionControls,visRowControl ,visCtlYCon,L"Controls.{r}.YCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",L"{r}",L"YDyn",visSectionControls,visRowControl ,visCtlYDyn,L"Controls.{r}.YDyn",L"",L""},
	{L"Document Properties",L"",L"AddMarkup",visSectionObject,visRowDoc,visDocAddMarkup,L"AddMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",L"",L"DocLangID",visSectionObject,visRowDoc,visDocLangID,L"DocLangID",L"",L""},
	{L"Document Properties",L"",L"DocLockDuplicatePage",visSectionObject,visRowDoc,visDocLockDuplicatePage,L"DocLockDuplicatePage",L"",L""},
	{L"Document Properties",L"",L"DocLocReplace",visSectionObject,visRowDoc,visDocLockReplace,L"DocLocReplace",L"",L""},
	{L"Document Properties",L"",L"LockPreview",visSectionObject,visRowDoc,visDocLockPreview,L"LockPreview",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",L"",L"NoCoauth",visSectionObject,visRowDoc,visDocNoCoauth,L"NoCoauth",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",L"",L"OutputFormat",visSectionObject,visRowDoc,visDocOutputFormat,L"OutputFormat",L"ENUM",L"0;1;2"},
	{L"Document Properties",L"",L"PreviewQuality",visSectionObject,visRowDoc,visDocPreviewQuality,L"PreviewQuality",L"ENUM",L"0=visDocPreviewQualityDraft;1=visDocPreviewQualityDetailed"},
	{L"Document Properties",L"",L"PreviewScope",visSectionObject,visRowDoc,visDocPreviewScope,L"PreviewScope",L"ENUM",L"0=visDocPreviewScope1stPage;1=visDocPreviewScopeNone;2=visDocPreviewScopeAllPages"},
	{L"Document Properties",L"",L"ViewMarkup",visSectionObject,visRowDoc,visDocViewMarkup,L"ViewMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Events",L"",L"EventDblClick",visSectionObject,visRowEvent,visEvtCellDblClick,L"EventDblClick",L"",L""},
	{L"Events",L"",L"EventDrop",visSectionObject,visRowEvent,visEvtCellDrop,L"EventDrop",L"",L""},
	{L"Events",L"",L"EventMultiDrop",visSectionObject,visRowEvent,visEvtCellMultiDrop,L"EventMultiDrop",L"",L""},
	{L"Events",L"",L"EventXFMod",visSectionObject,visRowEvent,visEvtCellXFMod,L"EventXFMod",L"",L""},
	{L"Events",L"",L"TheText",visSectionObject,visRowEvent,visEvtCellTheText,L"TheText",L"",L""},
	{L"Fill Format",L"",L"FillBkgnd",visSectionObject,visRowFill,visFillBkgnd,L"FillBkgnd",L"",L""},
	{L"Fill Format",L"",L"FillBkgndTrans",visSectionObject,visRowFill,visFillBkgndTrans,L"FillBkgndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",L"",L"FillForegnd",visSectionObject,visRowFill,visFillForegnd,L"FillForegnd",L"",L""},
	{L"Fill Format",L"",L"FillForegndTrans",visSectionObject,visRowFill,visFillForegndTrans,L"FillForegndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",L"",L"FillPattern",visSectionObject,visRowFill,visFillPattern,L"FillPattern",L"ENUM",L"0;1;2 - 40"},
	{L"Fill Format",L"",L"ShapeShdwBlur",visSectionObject,visRowFill,visFillShdwBlur,L"ShapeShdwBlur",L"",L""},
	{L"Fill Format",L"",L"ShapeShdwObliqueAngle",visSectionObject,visRowFill,visFillShdwObliqueAngle,L"ShapeShdwObliqueAngle",L"",L""},
	{L"Fill Format",L"",L"ShapeShdwOffsetX",visSectionObject,visRowFill,visFillShdwOffsetX,L"ShapeShdwOffsetX",L"",L""},
	{L"Fill Format",L"",L"ShapeShdwOffsetY",visSectionObject,visRowFill,visFillShdwOffsetY,L"ShapeShdwOffsetY",L"",L""},
	{L"Fill Format",L"",L"ShapeShdwScaleFactor",visSectionObject,visRowFill,visFillShdwScaleFactor,L"ShapeShdwScaleFactor",L"",L""},
	{L"Fill Format",L"",L"ShapeShdwShow",visSectionObject,visRowFill,visFillShdwShow,L"ShapeShdwShow",L"ENUM",L"0;1;2"},
	{L"Fill Format",L"",L"ShapeShdwType",visSectionObject,visRowFill,visFillShdwType,L"ShapeShdwType",L"ENUM",L"0=visFSTPageDefault;1=visFSTSimple;2=visFSTOblique"},
	{L"Fill Format",L"",L"ShdwForegnd",visSectionObject,visRowFill,visFillShdwForegnd,L"ShdwForegnd",L"",L""},
	{L"Fill Format",L"",L"ShdwForegndTrans",visSectionObject,visRowFill,visFillShdwForegndTrans,L"ShdwForegndTrans",L"",L"0 - 100"},
	{L"Fill Format",L"",L"ShdwPattern",visSectionObject,visRowFill,visFillShdwPattern,L"ShdwPattern",L"",L"0;1;2 - 40"},
	{L"Foreign Image Info",L"",L"ClippingPath",visSectionObject,visRowForeign,visFrgnImgClippingPath,L"ClippingPath",L"",L""},
	{L"Foreign Image Info",L"",L"ImgHeight",visSectionObject,visRowForeign,visFrgnImgHeight,L"ImgHeight",L"",L""},
	{L"Foreign Image Info",L"",L"ImgOffsetX",visSectionObject,visRowForeign,visFrgnImgOffsetX,L"ImgOffsetX",L"",L""},
	{L"Foreign Image Info",L"",L"ImgOffsetY",visSectionObject,visRowForeign,visFrgnImgOffsetY,L"ImgOffsetY",L"",L""},
	{L"Foreign Image Info",L"",L"ImgWidth",visSectionObject,visRowForeign,visFrgnImgWidth,L"ImgWidth",L"",L""},
	{L"Geometry",L"{i}",L"NoFill",visSectionFirstComponent,visRowComponent,visCompNoFill,L"Geometry{i}.NoFill",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",L"{i}",L"NoLine",visSectionFirstComponent,visRowComponent,visCompNoLine,L"Geometry{i}.NoLine",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",L"{i}",L"NoQuickDrag",visSectionFirstComponent,visRowComponent,visCompNoQuickDrag,L"Geometry{i}.NoQuickDrag",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",L"{i}",L"NoShow",visSectionFirstComponent,visRowComponent,visCompNoShow,L"Geometry{i}.NoShow",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",L"{i}",L"NoSnap",visSectionFirstComponent,visRowComponent,visCompNoSnap,L"Geometry{i}.NoSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",L"{i}",L"X",visSectionFirstComponent,visRowVertex ,visX,L"Geometry{i}.X{j}",L"",L""},
	{L"Geometry",L"{i}",L"Y",visSectionFirstComponent,visRowVertex ,visY,L"Geometry{i}.Y{j}",L"",L""},
	{L"Geometry",L"{i}",L"A",visSectionFirstComponent,visRowVertex ,2,L"Geometry{i}.A{j}",L"",L""},
	{L"Geometry",L"{i}",L"B",visSectionFirstComponent,visRowVertex ,3,L"Geometry{i}.B{j}",L"",L""},
	{L"Geometry",L"{i}",L"C",visSectionFirstComponent,visRowVertex ,4,L"Geometry{i}.C{j}",L"",L""},
	{L"Geometry",L"{i}",L"D",visSectionFirstComponent,visRowVertex ,5,L"Geometry{i}.D{j}",L"",L""},
	{L"Geometry",L"{i}",L"E",visSectionFirstComponent,visRowVertex ,6,L"Geometry{i}.E{j}",L"",L""},
	{L"Glue Info",L"",L"BegTrigger",visSectionObject,visRowGroup,visBegTrigger,L"BegTrigger",L"",L""},
	{L"Glue Info",L"",L"EndTrigger",visSectionObject,visRowMisc,visEndTrigger,L"EndTrigger",L"",L""},
	{L"Glue Info",L"",L"GlueType",visSectionObject,visRowMisc,visGlueType,L"GlueType",L"ENUM",L"0=visGlueTypeDefault;2=visGlueTypeWalking;4=visGlueTypeNoWalking;8=visGlueTypeNoWalkingTo"},
	{L"Glue Info",L"",L"WalkPreference",visSectionObject,visRowMisc,visWalkPref,L"WalkPreference",L"ENUM",L"1=visWalkPrefBegNS;2=visWalkPrefEndNS"},
	{L"Gradient Properties",L"",L"FillGradientAngle",visSectionObject,visRowGradientProperties,visFillGradientAngle,L"FillGradientAngle",L"",L""},
	{L"Gradient Properties",L"",L"FillGradientDir",visSectionObject,visRowGradientProperties,visFillGradientDir,L"FillGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",L"",L"FillGradientEnabled",visSectionObject,visRowGradientProperties,visFillGradientEnabled,L"FillGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",L"",L"LineGradientAngle",visSectionObject,visRowGradientProperties,visLineGradientAngle,L"LineGradientAngle",L"",L""},
	{L"Gradient Properties",L"",L"LineGradientDir",visSectionObject,visRowGradientProperties,visLineGradientDir,L"LineGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",L"",L"LineGradientEnabled",visSectionObject,visRowGradientProperties,visLineGradientEnabled,L"LineGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",L"",L"RotateGradientWithShape",visSectionObject,visRowGradientProperties,visRotateGradientWithShape,L"RotateGradientWithShape",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",L"",L"UseGroupGradient",visSectionObject,visRowGradientProperties,visUseGroupGradient,L"UseGroupGradient",L"",L""},
	{L"Group Properties",L"",L"DisplayMode",visSectionObject,visRowGroup,visGroupDisplayMode,L"DisplayMode",L"ENUM",L"0=visGrpDispModeNone;1=visGrpDispModeBack;2=visGrpDispModeFront"},
	{L"Group Properties",L"",L"DontMoveChildren",visSectionObject,visRowGroup,visGroupDontMoveChildren,L"DontMoveChildren",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",L"",L"IsDropTarget",visSectionObject,visRowGroup,visGroupIsDropTarget,L"IsDropTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",L"",L"IsSnapTarget",visSectionObject,visRowGroup,visGroupIsSnapTarget,L"IsSnapTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",L"",L"IsTextEditTarget",visSectionObject,visRowGroup,visGroupIsTextEditTarget,L"IsTextEditTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",L"",L"SelectMode",visSectionObject,visRowGroup,visGroupSelectMode,L"SelectMode",L"",L"0=visGrpSelModeGroupOnly;1=visGrpSelModeGroup1st;2=visGrpSelModeMembers1st"},
	{L"Hyperlinks",L"{r}",L"Address",visSectionHyperlink,visRow1stHyperlink ,visHLinkAddress,L"Hyperlink.{r}.Address",L"",L""},
	{L"Hyperlinks",L"{r}",L"Default",visSectionHyperlink,visRow1stHyperlink ,visHLinkDefault,L"Hyperlink.{r}.Default",L"",L""},
	{L"Hyperlinks",L"{r}",L"Description",visSectionHyperlink,visRow1stHyperlink ,visHLinkDescription,L"Hyperlink.{r}.Description",L"",L""},
	{L"Hyperlinks",L"{r}",L"ExtraInfo",visSectionHyperlink,visRow1stHyperlink ,visHLinkExtraInfo,L"Hyperlink.{r}.ExtraInfo",L"",L""},
	{L"Hyperlinks",L"{r}",L"Frame",visSectionHyperlink,visRow1stHyperlink ,visHLinkFrame,L"Hyperlink.{r}.Frame",L"",L""},
	{L"Hyperlinks",L"{r}",L"Invisible",visSectionHyperlink,visRow1stHyperlink ,visHLinkInvisible,L"Hyperlink.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",L"{r}",L"NewWindow",visSectionHyperlink,visRow1stHyperlink ,visHLinkNewWin,L"Hyperlink.{r}.NewWindow",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",L"{r}",L"SortKey",visSectionHyperlink,visRow1stHyperlink ,visHLinkSortKey,L"Hyperlink.{r}.SortKey",L"",L""},
	{L"Hyperlinks",L"{r}",L"SubAddress",visSectionHyperlink,visRow1stHyperlink ,visHLinkSubAddress,L"Hyperlink.{r}.SubAddress",L"",L""},
	{L"Image Properties",L"",L"Blur",visSectionObject,visRowImage,visImageBlur,L"Blur",L"",L""},
	{L"Image Properties",L"",L"Brightness",visSectionObject,visRowImage,visImageBrightness,L"Brightness",L"",L""},
	{L"Image Properties",L"",L"Contrast",visSectionObject,visRowImage,visImageContrast,L"Contrast",L"",L""},
	{L"Image Properties",L"",L"Denoise",visSectionObject,visRowImage,visImageDenoise,L"Denoise",L"",L""},
	{L"Image Properties",L"",L"Gamma",visSectionObject,visRowImage,visImageGamma,L"Gamma",L"",L""},
	{L"Image Properties",L"",L"Sharpen",visSectionObject,visRowImage,visImageSharpen,L"Sharpen",L"",L""},
	{L"Image Properties",L"",L"Transparency",visSectionObject,visRowImage,visImageTransparency,L"Transparency",L"",L"0 - 100"},
	{L"Layers",L"{i}",L"Active",visSectionLayer,visRowLayer ,visLayerActive,L"Layers.Active[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",L"{i}",L"Color",visSectionLayer,visRowLayer ,visLayerColor,L"Layers.Color[{i}]",L"",L""},
	{L"Layers",L"{i}",L"ColorTrans",visSectionLayer,visRowLayer ,visLayerColorTrans,L"Layers.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Layers",L"{i}",L"Glue",visSectionLayer,visRowLayer ,visLayerGlue,L"Layers.Glue[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",L"{i}",L"Locked",visSectionLayer,visRowLayer ,visLayerLock,L"Layers.Locked[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",L"{i}",L"Print",visSectionLayer,visRowLayer ,visDocPreviewScope,L"Layers.Print[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",L"{i}",L"Snap",visSectionLayer,visRowLayer ,visLayerSnap,L"Layers.Snap[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",L"{i}",L"Visible",visSectionLayer,visRowLayer ,visLayerVisible,L"Layers.Visible[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Line Format",L"",L"BeginArrow",visSectionObject,visRowLine,visLineBeginArrow,L"BeginArrow",L"",L"0;1 - 45"},
	{L"Line Format",L"",L"BeginArrowSize",visSectionObject,visRowLine,visLineBeginArrowSize,L"BeginArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",L"",L"CompoundType",visSectionObject,visRowLine,visCompoundType,L"CompoundType",L"",L"0;1;2;3;4"},
	{L"Line Format",L"",L"EndArrow",visSectionObject,visRowLine,visLineEndArrow,L"EndArrow",L"",L"0;1 - 45"},
	{L"Line Format",L"",L"EndArrowSize",visSectionObject,visRowLine,visLineEndArrowSize,L"EndArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",L"",L"LineCap",visSectionObject,visRowLine,visLineEndCap,L"LineCap",L"",L"0;1;2"},
	{L"Line Format",L"",L"LineColor",visSectionObject,visRowLine,visLineColor,L"LineColor",L"",L""},
	{L"Line Format",L"",L"LineColorTrans",visSectionObject,visRowLine,visLineColorTrans,L"LineColorTrans",L"",L"0 - 100"},
	{L"Line Format",L"",L"LinePattern",visSectionObject,visRowLine,visLinePattern,L"LinePattern",L"",L"0;1;2 - 23"},
	{L"Line Format",L"",L"LineWeight",visSectionObject,visRowLine,visLineWeight,L"LineWeight",L"",L""},
	{L"Line Format",L"",L"Rounding",visSectionObject,visRowLine,visLineRounding,L"Rounding",L"",L""},
	{L"Miscellaneous",L"",L"Calendar",visSectionObject,visRowMisc,visObjCalendar,L"Calendar",L"",L""},
	{L"Miscellaneous",L"",L"Comment",visSectionObject,visRowMisc,visComment,L"Comment",L"",L""},
	{L"Miscellaneous",L"",L"DropOnPageScale",visSectionObject,visRowMisc,visObjDropOnPageScale,L"DropOnPageScale",L"",L""},
	{L"Miscellaneous",L"",L"DynFeedback",visSectionObject,visRowMisc,visDynFeedback,L"DynFeedback",L"",L"0=visDynFBDefault;1=visDynFBUCon3Leg;2=visDynFBUCon5Leg"},
	{L"Miscellaneous",L"",L"HideText",visSectionObject,visRowMisc,visHideText,L"HideText",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"IsDropSource",visSectionObject,visRowMisc,visDropSource,L"IsDropSource",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"LangID",visSectionObject,visRowMisc,visObjLangID,L"LangID",L"",L""},
	{L"Miscellaneous",L"",L"LocalizeMerge",visSectionObject,visRowMisc,visObjLocalizeMerge,L"LocalizeMerge",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NoAlignBox",visSectionObject,visRowMisc,visNoAlignBox,L"NoAlignBox",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NoCtlHandles",visSectionObject,visRowMisc,visNoCtlHandles,L"NoCtlHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NoLiveDynamics",visSectionObject,visRowMisc,visNoLiveDynamics,L"NoLiveDynamics",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NonPrinting",visSectionObject,visRowMisc,visNonPrinting,L"NonPrinting",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NoObjHandles",visSectionObject,visRowMisc,visNoObjHandles,L"NoObjHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"NoProofing",visSectionObject,visRowMisc,visObjNoProofing,L"NoProofing",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",L"",L"ObjType",visSectionObject,visRowMisc,visLOFlags,L"ObjType",L"",L"0=visLOFlagsVisDecides;1=visLOFlagsPlacable;2=visLOFlagsRoutable;4=visLOFlagsDont;8=visLOFlagsPNRGroup"},
	{L"Miscellaneous",L"",L"UpdateAlignBox",visSectionObject,visRowMisc,visUpdateAlignBox,L"UpdateAlignBox",L"",L""},
	{L"Page Layout",L"",L"AvenueSizeY",visSectionObject,visRowPageLayout,visPLOAvenueSizeX,L"AvenueSizeY",L"",L""},
	{L"Page Layout",L"",L"AvenueSizeY",visSectionObject,visRowPageLayout,visPLOAvenueSizeY,L"AvenueSizeY",L"",L""},
	{L"Page Layout",L"",L"AvoidPageBreaks",visSectionObject,visRowPageLayout,visPLOAvoidPageBreaks,L"AvoidPageBreaks",L"",L""},
	{L"Page Layout",L"",L"BlockSizeX",visSectionObject,visRowPageLayout,visPLOBlockSizeX,L"BlockSizeX",L"",L""},
	{L"Page Layout",L"",L"BlockSizeY",visSectionObject,visRowPageLayout,visPLOBlockSizeY,L"BlockSizeY",L"",L""},
	{L"Page Layout",L"",L"CtrlAsInput",visSectionObject,visRowPageLayout,visPLOCtrlAsInput,L"CtrlAsInput",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",L"",L"DynamicsOff",visSectionObject,visRowPageLayout,visPLODynamicsOff,L"DynamicsOff",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",L"",L"EnableGrid",visSectionObject,visRowPageLayout,visPLOEnableGrid,L"EnableGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",L"",L"LineAdjustFrom",visSectionObject,visRowPageLayout,visPLOLineAdjustFrom,L"LineAdjustFrom",L"",L"0=visPLOLineAdjustFromNotRelated;1=visPLOLineAdjustFromAll;2=visPLOLineAdjustFromNone;3=visPLOLineAdjustFromRoutingDefault"},
	{L"Page Layout",L"",L"LineAdjustTo",visSectionObject,visRowPageLayout,visPLOLineAdjustTo,L"LineAdjustTo",L"",L"0=visPLOLineAdjustToDefault;1=visPLOLineAdjustToAll;2=visPLOLineAdjustToNone;3=visPLOLineAdjustToRelated"},
	{L"Page Layout",L"",L"LineJumpCode",visSectionObject,visRowPageLayout,visPLOJumpCode,L"LineJumpCode",L"",L"0=visPLOJumpNone;1=visPLOJumpHorizontal;2=visPLOJumpVertical;3=visPLOJumpLastRouted;4=visPLOJumpDisplayOrder;5=visPLOJumpReverseDisplayOrder"},
	{L"Page Layout",L"",L"LineJumpFactorX",visSectionObject,visRowPageLayout,visPLOJumpFactorX,L"LineJumpFactorX",L"",L""},
	{L"Page Layout",L"",L"LineJumpFactorY",visSectionObject,visRowPageLayout,visPLOJumpFactorY,L"LineJumpFactorY",L"",L""},
	{L"Page Layout",L"",L"LineJumpStyle",visSectionObject,visRowPageLayout,visPLOJumpStyle,L"LineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Page Layout",L"",L"LineRouteExt",visSectionObject,visRowPageLayout,visPLOLineRouteExt,L"LineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Page Layout",L"",L"LineToLineX",visSectionObject,visRowPageLayout,visPLOLineToLineX,L"LineToLineX",L"",L""},
	{L"Page Layout",L"",L"LineToLineY",visSectionObject,visRowPageLayout,visPLOLineToLineY,L"LineToLineY",L"",L""},
	{L"Page Layout",L"",L"LineToNodeX",visSectionObject,visRowPageLayout,visPLOLineToNodeX,L"LineToNodeX",L"",L""},
	{L"Page Layout",L"",L"LineToNodeY",visSectionObject,visRowPageLayout,visPLOLineToNodeY,L"LineToNodeY",L"",L""},
	{L"Page Layout",L"",L"PageLineJumpDirX",visSectionObject,visRowPageLayout,visPLOJumpDirX,L"PageLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Page Layout",L"",L"PageLineJumpDirY",visSectionObject,visRowPageLayout,visPLOJumpDirY,L"PageLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Page Layout",L"",L"PageShapeSplit",visSectionObject,visRowPageLayout,visPLOSplit,L"PageShapeSplit",L"",L"0=visPLOSplitNone;1=visPLOSplitAllow"},
	{L"Page Layout",L"",L"PlaceDepth",visSectionObject,visRowPageLayout,visPLOPlaceDepth,L"PlaceDepth",L"",L"0=visPLOPlaceDepthDefault;1=visPLOPlaceDepthMedium;2=visPLOPlaceDepthDeep;3=visPLOPlaceDepthShallow"},
	{L"Page Layout",L"",L"PlaceFlip",visSectionObject,visRowPageLayout,visPLOPlaceFlip,L"PlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Page Layout",L"",L"PlaceStyle",visSectionObject,visRowPageLayout,visPLOPlaceStyle,L"PlaceStyle",L"",L""},
	{L"Page Layout",L"",L"PlowCode",visSectionObject,visRowPageLayout,visPLOPlowCode,L"PlowCode",L"",L"0=visPLOPlowNone;1=visPLOPlowAll"},
	{L"Page Layout",L"",L"ResizePage",visSectionObject,visRowPageLayout,visPLOResizePage,L"ResizePage",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",L"",L"RouteStyle",visSectionObject,visRowPageLayout,visPLORouteStyle,L"RouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Page Properties",L"",L"DrawingResizeType",visSectionObject,visRowPage,visPageDrawResizeType,L"DrawingResizeType",L"",L""},
	{L"Page Properties",L"",L"DrawingScale",visSectionObject,visRowPage,visPageDrawingScale,L"DrawingScale",L"",L""},
	{L"Page Properties",L"",L"DrawingScaleType",visSectionObject,visRowPage,visPageDrawScaleType,L"DrawingScaleType",L"",L"0=visNoScale;1=visArchitectural;2=visEngineering;3=visScaleCustom;4=visScaleMetric;5=visScaleMechanical"},
	{L"Page Properties",L"",L"DrawingSizeType",visSectionObject,visRowPage,visPageDrawSizeType,L"DrawingSizeType",L"",L"0=visPrintSetup;1=visTight;2=visStandard;3=visCustom;4=visLogical;5=visDSMetric;6=visDSEngr;7=visDSArch"},
	{L"Page Properties",L"",L"InhibitSnap",visSectionObject,visRowPage,visPageInhibitSnap,L"InhibitSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",L"",L"PageHeight",visSectionObject,visRowPage,visPageHeight,L"PageHeight",L"",L""},
	{L"Page Properties",L"",L"PageLockDuplicate",visSectionObject,visRowPage,visPageLockDuplicate,L"PageLockDuplicate",L"",L""},
	{L"Page Properties",L"",L"PageLockReplace",visSectionObject,visRowPage,visPageLockReplace,L"PageLockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",L"",L"PageScale",visSectionObject,visRowPage,visPageScale,L"PageScale",L"",L""},
	{L"Page Properties",L"",L"PageWidth",visSectionObject,visRowPage,visPageWidth,L"PageWidth",L"",L""},
	{L"Page Properties",L"",L"ShdwObliqueAngle",visSectionObject,visRowPage,visPageShdwObliqueAngle,L"ShdwObliqueAngle",L"",L""},
	{L"Page Properties",L"",L"ShdwOffsetX",visSectionObject,visRowPage,visPageShdwOffsetX,L"ShdwOffsetX",L"",L""},
	{L"Page Properties",L"",L"ShdwOffsetY",visSectionObject,visRowPage,visPageShdwOffsetY,L"ShdwOffsetY",L"",L""},
	{L"Page Properties",L"",L"ShdwScaleFactor",visSectionObject,visRowPage,visPageShdwScaleFactor,L"ShdwScaleFactor",L"",L""},
	{L"Page Properties",L"",L"ShdwType",visSectionObject,visRowPage,visPageShdwType,L"ShdwType",L"",L"1=visFSTSimple;2=visFSTOblique;3=visFSTInner"},
	{L"Page Properties",L"",L"UIVisibility",visSectionObject,visRowPage,visPageUIVisibility,L"UIVisibility",L"",L"0=visUIVNormal;1=visUIVHidden"},
	{L"Paragraph",L"{i}",L"Bullet",visSectionParagraph,visRowParagraph ,visBulletIndex,L"Para.Bullet[{i}]",L"",L"0;1;2;3;4;5;6;7"},
	{L"Paragraph",L"{i}",L"BulletFont",visSectionParagraph,visRowParagraph ,visBulletFont,L"Para.BulletFont[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"BulletFontSize",visSectionParagraph,visRowParagraph ,visBulletFontSize,L"Para.BulletFontSize[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"BulletStr",visSectionParagraph,visRowParagraph ,visBulletString,L"Para.BulletStr[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"Flags",visSectionParagraph,visRowParagraph ,visFlags,L"Para.Flags[{i}]",L"",L"0;1"},
	{L"Paragraph",L"{i}",L"HorzAlign",visSectionParagraph,visRowParagraph ,visHorzAlign,L"Para.HorzAlign[{i}]",L"",L"0=visHorzLeft;1=visHorzCenter;2=visHorzRight;3=visHorzJustify;4=visHorzForce"},
	{L"Paragraph",L"{i}",L"IndFirst",visSectionParagraph,visRowParagraph ,visIndentFirst,L"Para.IndFirst[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"IndLeft",visSectionParagraph,visRowParagraph ,visIndentLeft,L"Para.IndLeft[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"IndRight",visSectionParagraph,visRowParagraph ,visIndentRight,L"Para.IndRight[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"SpAfter",visSectionParagraph,visRowParagraph ,visSpaceAfter,L"Para.SpAfter[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"SpBefore",visSectionParagraph,visRowParagraph ,visSpaceBefore,L"Para.SpBefore[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"SpLine",visSectionParagraph,visRowParagraph ,visSpaceLine,L"Para.SpLine[{i}]",L"",L""},
	{L"Paragraph",L"{i}",L"TextPosAfterBullet",visSectionParagraph,visRowParagraph ,visTextPosAfterBullet,L"Para.TextPosAfterBullet[{i}]",L"",L""},
	{L"Print Properties",L"",L"CenterX",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterX,L"CenterX",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",L"",L"CenterY",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterY,L"CenterY",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",L"",L"OnPage",visSectionObject,visRowPrintProperties,visPrintPropertiesOnPage,L"OnPage",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",L"",L"PageBottomMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesBottomMargin,L"PageBottomMargin",L"",L""},
	{L"Print Properties",L"",L"PageLeftMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesLeftMargin,L"PageLeftMargin",L"",L""},
	{L"Print Properties",L"",L"PageRightMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesRightMargin,L"PageRightMargin",L"",L""},
	{L"Print Properties",L"",L"PagesX",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesX,L"PagesX",L"",L""},
	{L"Print Properties",L"",L"PagesY",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesY,L"PagesY",L"",L""},
	{L"Print Properties",L"",L"PageTopMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesTopMargin,L"PageTopMargin",L"",L""},
	{L"Print Properties",L"",L"PaperKind",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperKind,L"PaperKind",L"",L""},
	{L"Print Properties",L"",L"PrintGrid",visSectionObject,visRowPrintProperties,visPrintPropertiesPrintGrid,L"PrintGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",L"",L"PrintPageOrientation",visSectionObject,visRowPrintProperties,visPrintPropertiesPageOrientation,L"PrintPageOrientation",L"",L"0=visPPOSameAsPrinter;1=visPPOPortrait;2=visPPOLandscape"},
	{L"Print Properties",L"",L"ScaleX",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleX,L"ScaleX",L"",L""},
	{L"Print Properties",L"",L"ScaleY",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleY,L"ScaleY",L"",L""},
	{L"PrintProperties",L"",L"PaperSource",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperSource,L"PaperSource",L"",L""},
	{L"Protection",L"",L"LockAspect",visSectionObject,visRowLock,visLockAspect,L"LockAspect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockBegin",visSectionObject,visRowLock,visLockBegin,L"LockBegin",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockCalcWH",visSectionObject,visRowLock,visLockCalcWH,L"LockCalcWH",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockCrop",visSectionObject,visRowLock,visLockCrop,L"LockCrop",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockCustProp",visSectionObject,visRowLock,visLockCustProp,L"LockCustProp",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockDelete",visSectionObject,visRowLock,visLockDelete,L"LockDelete",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockEnd",visSectionObject,visRowLock,visLockEnd,L"LockEnd",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockFormat",visSectionObject,visRowLock,visLockFormat,L"LockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockFromGroupFormat",visSectionObject,visRowLock,visLockFromGroupFormat,L"LockFromGroupFormat",L"",L""},
	{L"Protection",L"",L"LockGroup",visSectionObject,visRowLock,visLockGroup,L"LockGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockHeight",visSectionObject,visRowLock,visLockHeight,L"LockHeight",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockMoveX",visSectionObject,visRowLock,visLockMoveX,L"LockMoveX",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockMoveY",visSectionObject,visRowLock,visLockMoveY,L"LockMoveY",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockReplace",visSectionObject,visRowLock,visLockReplace,L"LockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockRotate",visSectionObject,visRowLock,visLockRotate,L"LockRotate",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockSelect",visSectionObject,visRowLock,visLockSelect,L"LockSelect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockTextEdit",visSectionObject,visRowLock,visLockTextEdit,L"LockTextEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockThemeColors",visSectionObject,visRowLock,visLockThemeColors,L"LockThemeColors",L"",L""},
	{L"Protection",L"",L"LockThemeConnectors",visSectionObject,visRowLock,visLockThemeConnectors,L"LockThemeConnectors",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockThemeEffects",visSectionObject,visRowLock,visLockThemeEffects,L"LockThemeEffects",L"",L""},
	{L"Protection",L"",L"LockThemeFonts",visSectionObject,visRowLock,visLockThemeFonts,L"LockThemeFonts",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockThemeIndex",visSectionObject,visRowLock,visLockThemeIndex,L"LockThemeIndex",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockVariation",visSectionObject,visRowLock,visLockVariation,L"LockVariation",L"",L""},
	{L"Protection",L"",L"LockVtxEdit",visSectionObject,visRowLock,visLockVtxEdit,L"LockVtxEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",L"",L"LockWidth",visSectionObject,visRowLock,visLockWidth,L"LockWidth",L"BOOL",L"TRUE;FALSE"},
	{L"Quick Style",L"",L"QuickStyleEffectsMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleEffectsMatrix,L"QuickStyleEffectsMatrix",L"",L""},
	{L"Quick Style",L"",L"QuickStyleFillColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillColor,L"QuickStyleFillColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",L"",L"QuickStyleFillMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillMatrix,L"QuickStyleFillMatrix",L"",L""},
	{L"Quick Style",L"",L"QuickStyleFontColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontColor,L"QuickStyleFontColor",L"",L"0;1"},
	{L"Quick Style",L"",L"QuickStyleFontMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontMatrix,L"QuickStyleFontMatrix",L"",L""},
	{L"Quick Style",L"",L"QuickStyleLineColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineColor,L"QuickStyleLineColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",L"",L"QuickStyleLineMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineMatrix,L"QuickStyleLineMatrix",L"",L""},
	{L"Quick Style",L"",L"QuickStyleShadowColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleShadowColor,L"QuickStyleShadowColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",L"",L"QuickStyleType",visSectionObject,visRowQuickStyleProperties,visQuickStyleType,L"QuickStyleType",L"",L"0;1;2;3"},
	{L"Quick Style",L"",L"QuickStyleVariation",visSectionObject,visRowQuickStyleProperties,visQuickStyleVariation,L"QuickStyleVariation",L"",L"0;1;2;4;8"},
	{L"Reviewer",L"{i}",L"Color",visSectionReviewer,visRowReviewer ,visReviewerColor,L"Reviewer.Color[{i}]",L"",L""},
	{L"Reviewer",L"{i}",L"Initials",visSectionReviewer,visRowReviewer ,visReviewerInitials,L"Reviewer.Initials[{i}]",L"",L""},
	{L"Reviewer",L"{i}",L"Name",visSectionReviewer,visRowReviewer ,visReviewerName,L"Reviewer.Name[{i}]",L"",L""},
	{L"Ruler & Grid",L"",L"XGridDensity",visSectionObject,visRowRulerGrid,visXGridDensity,L"XGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",L"",L"XGridOrigin",visSectionObject,visRowRulerGrid,visXGridOrigin,L"XGridOrigin",L"",L""},
	{L"Ruler & Grid",L"",L"XGridSpacing",visSectionObject,visRowRulerGrid,visXGridSpacing,L"XGridSpacing",L"",L""},
	{L"Ruler & Grid",L"",L"XRulerDensity",visSectionObject,visRowRulerGrid,visXRulerDensity,L"XRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",L"",L"XRulerOrigin",visSectionObject,visRowRulerGrid,visXRulerOrigin,L"XRulerOrigin",L"",L""},
	{L"Ruler & Grid",L"",L"YGridDensity",visSectionObject,visRowRulerGrid,visYGridDensity,L"YGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",L"",L"YGridOrigin",visSectionObject,visRowRulerGrid,visYGridOrigin,L"YGridOrigin",L"",L""},
	{L"Ruler & Grid",L"",L"YGridSpacing",visSectionObject,visRowRulerGrid,visYGridSpacing,L"YGridSpacing",L"",L""},
	{L"Ruler & Grid",L"",L"YRulerDensity",visSectionObject,visRowRulerGrid,visYRulerDensity,L"YRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",L"",L"YRulerOrigin",visSectionObject,visRowRulerGrid,visYRulerOrigin,L"YRulerOrigin",L"",L""},
	{L"Shape Data",L"{r}",L"Ask",visSectionProp,visRowProp ,visCustPropsAsk,L"Prop.{r}.Ask",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",L"{r}",L"Calendar",visSectionProp,visRowProp ,visCustPropsCalendar,L"Prop.{r}.Calendar",L"",L""},
	{L"Shape Data",L"{r}",L"Format",visSectionProp,visRowProp ,visCustPropsFormat,L"Prop.{r}.Format",L"",L""},
	{L"Shape Data",L"{r}",L"Invisible",visSectionProp,visRowProp ,visCustPropsInvis,L"Prop.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",L"{r}",L"Label",visSectionProp,visRowProp ,visCustPropsLabel,L"Prop.{r}.Label",L"",L""},
	{L"Shape Data",L"{r}",L"LangID",visSectionProp,visRowProp ,visCustPropsLangID,L"Prop.{r}.LangID",L"",L""},
	{L"Shape Data",L"{r}",L"Prompt",visSectionProp,visRowProp ,visCustPropsPrompt,L"Prop.{r}.Prompt",L"",L""},
	{L"Shape Data",L"{r}",L"SortKey",visSectionProp,visRowProp ,visCustPropsSortKey,L"Prop.{r}.SortKey",L"",L""},
	{L"Shape Data",L"{r}",L"Type",visSectionProp,visRowProp ,visCustPropsType,L"Prop.{r}.Type",L"",L"0=visPropTypeString;1=visPropTypeListFix;2=visPropTypeNumber;3=visPropTypeBool;4=visPropTypeListVar;5=visPropTypeDate;6=visPropTypeDuration;7=visPropTypeCurrency"},
	{L"Shape Data",L"{r}",L"Value",visSectionProp,visRowProp ,visCustPropsValue,L"Prop.{r}.Value",L"",L""},
	{L"Shape Layout",L"",L"ConFixedCode",visSectionObject,visRowShapeLayout,visSLOConFixedCode,L"ConFixedCode",L"",L"0=visSLOConFixedRerouteFreely;1=visSLOConFixedRerouteAsNeeded;2=visSLOConFixedRerouteNever;3=visSLOConFixedRerouteOnCrossover;4=visSLOConFixedByAlgFrom;5=visSLOConFixedByAlgTo;6=visSLOConFixedByAlgFromTo"},
	{L"Shape Layout",L"",L"ConLineJumpCode",visSectionObject,visRowShapeLayout,visSLOJumpCode,L"ConLineJumpCode",L"",L"0=visSLOJumpDefault;1=visSLOJumpNever;2=visSLOJumpAlways;3=visSLOJumpOther;4=visSLOJumpNeither"},
	{L"Shape Layout",L"",L"ConLineJumpDirX",visSectionObject,visRowShapeLayout,visSLOJumpDirX,L"ConLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Shape Layout",L"",L"ConLineJumpDirY",visSectionObject,visRowShapeLayout,visSLOJumpDirY,L"ConLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Shape Layout",L"",L"ConLineJumpStyle",visSectionObject,visRowShapeLayout,visSLOJumpStyle,L"ConLineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Shape Layout",L"",L"ConLineRouteExt",visSectionObject,visRowShapeLayout,visSLOLineRouteExt,L"ConLineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Shape Layout",L"",L"DisplayLevel",visSectionObject,visRowShapeLayout,visSLODisplayLevel,L"DisplayLevel",L"",L""},
	{L"Shape Layout",L"",L"Relationships",visSectionObject,visRowShapeLayout,visSLORelationships,L"Relationships",L"",L""},
	{L"Shape Layout",L"",L"ShapeFixedCode",visSectionObject,visRowShapeLayout,visSLOFixedCode,L"ShapeFixedCode",L"",L"1=visSLOFixedPlacement;2=visSLOFixedPlow;4=visSLOFixedPermeablePlow;32=visSLOFixedConnPtsIgnore;64=visSLOFixedConnPtsOnly;128=visSLOFixedNoFoldToShape"},
	{L"Shape Layout",L"",L"ShapePermeablePlace",visSectionObject,visRowShapeLayout,visSLOPermeablePlace,L"ShapePermeablePlace",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",L"",L"ShapePermeableX",visSectionObject,visRowShapeLayout,visSLOPermX,L"ShapePermeableX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",L"",L"ShapePermeableY",visSectionObject,visRowShapeLayout,visSLOPermY,L"ShapePermeableY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",L"",L"ShapePlaceFlip",visSectionObject,visRowShapeLayout,visSLOPlaceFlip,L"ShapePlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Shape Layout",L"",L"ShapePlaceStyle",visSectionObject,visRowShapeLayout,visSLOPlaceStyle,L"ShapePlaceStyle",L"",L"visLOPlaceBottomToTop;visLOPlaceCircular;visLOPlaceCompactDownLeft;visLOPlaceCompactDownRight;visLOPlaceCompactLeftDown;visLOPlaceCompactLeftUp;visLOPlaceCompactRightDown;visLOPlaceCompactRightUp;visLOPlaceCompactUpLeft;visLOPlaceCompactUpRight;visLOPlaceDefault;visLOPlaceHierarchyBottomToTopCenter;visLOPlaceHierarchyBottomToTopLeft;visLOPlaceHierarchyBottomToTopRight;visLOPlaceHierarchyLeftToRightBottom;visLOPlaceHierarchyLeftToRightMiddle;visLOPlaceHierarchyLeftToRightTop;visLOPlaceHierarchyRightToLeftBottom;visLOPlaceHierarchyRightToLeftMiddle;visLOPlaceHierarchyRightToLeftTop;visLOPlaceHierarchyTopToBottomCenter;visLOPlaceHierarchyTopToBottomLeft;visLOPlaceHierarchyTopToBottomRight;visLOPlaceLeftToRight;visLOPlaceParentDefault;visLOPlaceRadial;visLOPlaceRightToLeft;visLOPlaceTopToBottom"},
	{L"Shape Layout",L"",L"ShapePlowCode",visSectionObject,visRowShapeLayout,visSLOPlowCode,L"ShapePlowCode",L"",L"0=visSLOPlowDefault;1=visSLOPlowNever;2=visSLOPlowAlways"},
	{L"Shape Layout",L"",L"ShapeRouteStyle",visSectionObject,visRowShapeLayout,visSLORouteStyle,L"ShapeRouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Shape Layout",L"",L"ShapeSplit",visSectionObject,visRowShapeLayout,visSLOSplit,L"ShapeSplit",L"",L"0=visSLOSplitNone;1=visSLOSplitAllow"},
	{L"Shape Layout",L"",L"ShapeSplittable",visSectionObject,visRowShapeLayout,visSLOSplittable,L"ShapeSplittable",L"",L"0=visSLOSplittableNone;1=visSLOSplittableAllow"},
	{L"Shape Transform",L"",L"Angle",visSectionObject,visRowXFormOut,visXFormAngle,L"Angle",L"",L""},
	{L"Shape Transform",L"",L"FlipX",visSectionObject,visRowXFormOut,visXFormFlipX,L"FlipX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",L"",L"FlipY",visSectionObject,visRowXFormOut,visXFormFlipY,L"FlipY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",L"",L"Height",visSectionObject,visRowXFormOut,visXFormHeight,L"Height",L"",L""},
	{L"Shape Transform",L"",L"LocPinX",visSectionObject,visRowXFormOut,visXFormLocPinX,L"LocPinX",L"",L""},
	{L"Shape Transform",L"",L"LocPinY",visSectionObject,visRowXFormOut,visXFormLocPinY,L"LocPinY",L"",L""},
	{L"Shape Transform",L"",L"PinX",visSectionObject,visRowXFormOut,visXFormPinX,L"PinX",L"",L""},
	{L"Shape Transform",L"",L"PinY",visSectionObject,visRowXFormOut,visXFormPinY,L"PinY",L"",L""},
	{L"Shape Transform",L"",L"ResizeMode",visSectionObject,visRowXFormOut,visXFormResizeMode,L"ResizeMode",L"",L"0=visXFormResizeDontCare;1=visXFormResizeSpread;2=visXFormResizeScale"},
	{L"Shape Transform",L"",L"Width",visSectionObject,visRowXFormOut,visXFormWidth,L"Width",L"",L""},
	{L"Style Properties",L"",L"EnableFillProps",visSectionObject,visRowStyle,visStyleIncludesFill,L"EnableFillProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",L"",L"EnableLineProps",visSectionObject,visRowStyle,visStyleIncludesLine,L"EnableLineProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",L"",L"EnableTextProps",visSectionObject,visRowStyle,visStyleIncludesText,L"EnableTextProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",L"",L"HideForApply",visSectionObject,visRowStyle,visStyleHidden,L"HideForApply",L"BOOL",L"TRUE;FALSE"},
	{L"Tabs",L"Tabs.{i}{j}",L"Tabs.{i}{j}",visSectionTab,visRowTab , visTabAlign,L"Tabs.{i}{j}",L"",L"0=visTabStopLeft;1=visTabStopCenter;2=visTabStopRight;3=visTabStopDecimal;4=visTabStopComma"},
	{L"Tabs",L"Tabs.{i}{j}",L"Tabs.{i}{j}",visSectionTab,visRowTab , visTabPos,L"Tabs.{i}{j}",L"",L""},
	{L"Text Block Format",L"",L"BottomMargin",visSectionObject,visRowText,visTxtBlkBottomMargin,L"BottomMargin",L"",L""},
	{L"Text Block Format",L"",L"DefaultTabstop",visSectionObject,visRowText,visTxtBlkDefaultTabStop,L"DefaultTabstop",L"",L""},
	{L"Text Block Format",L"",L"LeftMargin",visSectionObject,visRowText,visTxtBlkLeftMargin,L"LeftMargin",L"",L""},
	{L"Text Block Format",L"",L"RightMargin",visSectionObject,visRowText,visTxtBlkRightMargin,L"RightMargin",L"",L""},
	{L"Text Block Format",L"",L"TextBkgnd",visSectionObject,visRowText,visTxtBlkBkgnd,L"TextBkgnd",L"",L""},
	{L"Text Block Format",L"",L"TextBkgndTrans",visSectionObject,visRowText,visTxtBlkBkgndTrans,L"TextBkgndTrans",L"",L"0 - 100"},
	{L"Text Block Format",L"",L"TextDirection",visSectionObject,visRowText,visTxtBlkDirection,L"TextDirection",L"",L"0=visTxtBlkLeftToRight;1=visTxtBlkTopToBottom"},
	{L"Text Block Format",L"",L"TopMargin",visSectionObject,visRowText,visTxtBlkTopMargin,L"TopMargin",L"",L""},
	{L"Text Block Format",L"",L"VerticalAlign",visSectionObject,visRowText,visTxtBlkVerticalAlign,L"VerticalAlign",L"",L"0=visVertTop;1=visVertMiddle;2=visVertBottom"},
	{L"Text Fields",L"{i}",L"Calendar",visSectionTextField,visRowField ,visFieldCalendar,L"Fields.Calendar[{i}]",L"",L""},
	{L"Text Fields",L"{i}",L"Format",visSectionTextField,visRowField ,visFieldFormat,L"Fields.Format[{i}]",L"",L""},
	{L"Text Fields",L"{i}",L"ObjectKind",visSectionTextField,visRowField ,visFieldObjectKind,L"Fields.ObjectKind[{i}]",L"",L"0=visTFOKStandard;1=visTFOKHorizontaInVertical"},
	{L"Text Fields",L"{i}",L"Type",visSectionTextField,visRowField ,visFieldType,L"Fields.Type[{i}]",L"",L"0;2;5;6;7"},
	{L"Text Fields",L"{i}",L"UICat",visSectionTextField,visRowField ,visFieldUICategory,L"Fields.UICat[{i}]",L"",L""},
	{L"Text Fields",L"{i}",L"UICod",visSectionTextField,visRowField ,visFieldUICode,L"Fields.UICod[{i}]",L"",L""},
	{L"Text Fields",L"{i}",L"UIFmt",visSectionTextField,visRowField ,visFieldUIFormat,L"Fields.UIFmt[{i}]",L"",L""},
	{L"Text Fields",L"{i}",L"Value",visSectionTextField,visRowField ,visFieldCell,L"Fields.Value[{i}]",L"",L""},
	{L"Text Transform",L"",L"TxtAngle",visSectionObject,visRowTextXForm,visXFormAngle,L"TxtAngle",L"",L""},
	{L"Text Transform",L"",L"TxtHeight",visSectionObject,visRowTextXForm,visXFormHeight,L"TxtHeight",L"",L""},
	{L"Text Transform",L"",L"TxtLocPinX",visSectionObject,visRowTextXForm,visXFormLocPinX,L"TxtLocPinX",L"",L""},
	{L"Text Transform",L"",L"TxtLocPinY",visSectionObject,visRowTextXForm,visXFormLocPinY,L"TxtLocPinY",L"",L""},
	{L"Text Transform",L"",L"TxtPinX",visSectionObject,visRowTextXForm,visXFormPinX,L"TxtPinX",L"",L""},
	{L"Text Transform",L"",L"TxtPinY",visSectionObject,visRowTextXForm,visXFormPinY,L"TxtPinY",L"",L""},
	{L"Text Transform",L"",L"TxtWidth",visSectionObject,visRowTextXForm,visXFormWidth,L"TxtWidth",L"",L""},
	{L"Theme Properties",L"",L"ColorSchemeIndex",visSectionObject,visRowThemeProperties,visColorSchemeIndex,L"ColorSchemeIndex",L"",L""},
	{L"Theme Properties",L"",L"ConnectorSchemeIndex",visSectionObject,visRowThemeProperties,visConnectorSchemeIndex,L"ConnectorSchemeIndex",L"",L""},
	{L"Theme Properties",L"",L"EffectSchemeIndex",visSectionObject,visRowThemeProperties,visEffectSchemeIndex,L"EffectSchemeIndex",L"",L""},
	{L"Theme Properties",L"",L"EmbellishmentIndex",visSectionObject,visRowThemeProperties,visEmbellishmentIndex,L"EmbellishmentIndex",L"",L""},
	{L"Theme Properties",L"",L"FontSchemeIndex",visSectionObject,visRowThemeProperties,visFontSchemeIndex,L"FontSchemeIndex",L"",L""},
	{L"Theme Properties",L"",L"ThemeIndex",visSectionObject,visRowThemeProperties,visThemeIndex,L"ThemeIndex",L"",L""},
	{L"Theme Properties",L"",L"VariationColorIndex",visSectionObject,visRowThemeProperties,visVariationColorIndex,L"VariationColorIndex",L"",L""},
	{L"Theme Properties",L"",L"VariationStyleIndex",visSectionObject,visRowThemeProperties,visVariationStyleIndex,L"VariationStyleIndex",L"",L""},
	{L"User-Defined Cells",L"{r}",L"Prompt",visSectionUser,visRowUser ,visUserPrompt,L"User.{r}.Prompt",L"",L""},
	{L"User-Defined Cells",L"{r}",L"Value",visSectionUser,visRowUser ,visUserValue,L"User.{r}.Value",L"",L""},
	{L"Scratch",L"{i}",L"X",visSectionScratch,visRowScratch,visScratchX,L"Scratch.X{i}",L"",L""},
	{L"Scratch",L"{i}",L"Y",visSectionScratch,visRowScratch,visScratchY,L"Scratch.Y{i}",L"",L""},
	{L"Scratch",L"{i}",L"A",visSectionScratch,visRowScratch,visScratchA,L"Scratch.A{i}",L"",L""},
	{L"Scratch",L"{i}",L"B",visSectionScratch,visRowScratch,visScratchB,L"Scratch.B{i}",L"",L""},
	{L"Scratch",L"{i}",L"C",visSectionScratch,visRowScratch,visScratchC,L"Scratch.C{i}",L"",L""},
	{L"Scratch",L"{i}",L"D",visSectionScratch,visRowScratch,visScratchD,L"Scratch.D{i}",L"",L""},

};

void AddNameMatchResult(const CString& mask, CString name, 
						short s, short r, short c, 
						const CString& s_name, const CString& r_name, const CString& c_name,
						std::vector<SRC> &result)
{
	Strings masks;
	SplitList(mask, L",", masks);

	for (Strings::const_iterator it = masks.begin(); it != masks.end(); ++it)
	{
		CString mask = *it;
		mask.Trim();

		if (StringIsLike(mask, name))
		{
			SRC src;

			src.name = name;
			src.s = s;
			src.s_name = s_name;
			src.r = r;
			src.r_name = r_name;
			src.c = c;
			src.c_name = c_name;

			result.push_back(src);
			return;
		}
	}
}

void GetVariableIndexedSectionCellNames(IVShapePtr shape, short section_no, const CString& mask, std::vector<SRC>& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	for (short r = 0; r < section->Count; ++r)
	{
		for (size_t i = 0; i < _countof(ss_info); ++i)
		{
			if (section_no != ss_info[i].s)
				continue;

			CString name = ss_info[i].name;

			CString r_name = FormatString(L"%d", 1 + r);
			name.Replace(L"{i}", r_name);

			AddNameMatchResult(mask, name, 
				section_no, r, ss_info[i].c, 
				ss_info[i].s_name, r_name, ss_info[i].c_name,
				result);
		}
	}
}

void GetVariableNamedSectionCellNames(IVShapePtr shape, short section_no, const CString& mask, std::vector<SRC>& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	for (short r = 0; r < section->Count; ++r)
 	{
		IVRowPtr row = section->GetRow(r);
		CString row_name = row->NameU;

		for (size_t i = 0; i < _countof(ss_info); ++i)
		{
			if (section_no != ss_info[i].s)
				continue;

			CString name = ss_info[i].name;
			name.Replace(L"{r}", row_name);

			AddNameMatchResult(mask, name, 
				section_no, r, ss_info[i].c, 
				ss_info[i].s_name, row_name, ss_info[i].c_name,
				result);
		}
	}
}

void GetVariableGeometrySectionCellNames(IVShapePtr shape, const CString& mask, std::vector<SRC>& result)
{
	for (short s = 0; s < shape->GeometryCount; ++s)
	{
		IVSectionPtr section = shape->GetSection(visSectionFirstComponent + s);

		CString s_name = FormatString(L"%d", 1 + s);

		for (short r = 0; r < section->Count; ++r)
		{
			CString r_name = FormatString(L"%d", 1 + r);

			for (size_t i = 0; i < _countof(ss_info); ++i)
			{
				if (ss_info[i].s != visSectionFirstComponent)
					continue;

				if (ss_info[i].r == visRowComponent)
				{
					if (r == 0)
					{
						CString name = ss_info[i].name;
						name.Replace(L"{i}", s_name);

						AddNameMatchResult(mask, name, 
							visSectionFirstComponent + s, r, ss_info[i].c, 
							s_name, L"", ss_info[i].c_name,
							result);
					}
					continue;
				}

				CString name = ss_info[i].name;
				name.Replace(L"{i}", s_name);
				name.Replace(L"{j}", r_name);

				AddNameMatchResult(mask, name, 
					visSectionFirstComponent + s, r, ss_info[i].c, 
					s_name, r_name, ss_info[i].c_name,
					result);
			}
		}
	}
}

void GetSimpleSectionCellNames(IVShapePtr shape, const CString& mask, std::vector<SRC>& result)
{
	for (size_t i = 0; i < _countof(ss_info); ++i)
	{
		if (ss_info[i].s != visSectionObject)
			continue;

		CString name = ss_info[i].name;

		AddNameMatchResult(mask, name, 
			ss_info[i].s, ss_info[i].r, ss_info[i].c, 
			ss_info[i].s_name, L"", ss_info[i].c_name,
			result);
	}
}

void GetCellNames(IVShapePtr shape, const CString& cell_name_mask, std::vector<SRC>& result)
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
	GetVariableIndexedSectionCellNames(shape, visSectionScratch, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionCharacter, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionConnectionPts, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionLayer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionParagraph, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionReviewer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionTextField, cell_name_mask, result);

	GetVariableGeometrySectionCellNames(shape, cell_name_mask, result);

	GetSimpleSectionCellNames(shape, cell_name_mask, result);
}
