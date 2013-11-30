
#include "stdafx.h"
#include "import/VISLIB.tlh"
#include "lib/Utils.h"
#include "Addin.h"
#include "ShapeSheet.h"

//TODO: tab support {i}{j}

namespace shapesheet {


// 
// This ShapeSheet data is copy-pasted from Excel file (in the "Data" folder)
// If you want to edit it, it is recommend to edit that Excel file instead and then copy-paste!
// 

SSInfo ss_global[] = {

	{0,L"1-D Endpoints",L"",L"BeginX",visSectionObject,visRowXForm1D,vis1DBeginX,L"BeginX",L""},
	{0,L"1-D Endpoints",L"",L"BeginY",visSectionObject,visRowXForm1D,vis1DBeginY,L"BeginY",L""},
	{0,L"1-D Endpoints",L"",L"EndX",visSectionObject,visRowXForm1D,vis1DEndX,L"EndX",L""},
	{0,L"1-D Endpoints",L"",L"EndY",visSectionObject,visRowXForm1D,vis1DEndY,L"EndY",L""},
	{15,L"3-D Rotation Properties",L"",L"RotationXAngle",visSectionObject,visRow3DRotationProperties,visRotationXAngle,L"RotationXAngle",L""},
	{15,L"3-D Rotation Properties",L"",L"RotationYAngle",visSectionObject,visRow3DRotationProperties,visRotationYAngle,L"RotationYAngle",L""},
	{15,L"3-D Rotation Properties",L"",L"RotationZAngle",visSectionObject,visRow3DRotationProperties,visRotationZAngle,L"RotationZAngle",L""},
	{15,L"3-D Rotation Properties",L"",L"RotationType",visSectionObject,visRow3DRotationProperties,visRotationType,L"RotationType",L"0-6"},
	{15,L"3-D Rotation Properties",L"",L"Perspective",visSectionObject,visRow3DRotationProperties,visPerspective,L"Perspective",L""},
	{15,L"3-D Rotation Properties",L"",L"DistanceFromGround",visSectionObject,visRow3DRotationProperties,visDistanceFromGround,L"DistanceFromGround",L""},
	{15,L"3-D Rotation Properties",L"",L"KeepTextFlat",visSectionObject,visRow3DRotationProperties,visKeepTextFlat,L"KeepTextFlat",L"TRUE;FALSE"},
	{11,L"Action Tags",L"{r}",L"X",visSectionSmartTag,visRowSmartTag ,visSmartTagX,L"SmartTags.{r}.X",L""},
	{11,L"Action Tags",L"{r}",L"Y",visSectionSmartTag,visRowSmartTag ,visSmartTagY,L"SmartTags.{r}.Y",L""},
	{11,L"Action Tags",L"{r}",L"TagName",visSectionSmartTag,visRowSmartTag ,visSmartTagName,L"SmartTags.{r}.TagName",L""},
	{11,L"Action Tags",L"{r}",L"XJustify",visSectionSmartTag,visRowSmartTag ,visSmartTagXJustify,L"SmartTags.{r}.XJustify",L"0=visSmartTagXJustifyLeft;1=visSmartTagXJustifyCenter;2=visSmartTagXJustifyRight"},
	{11,L"Action Tags",L"{r}",L"YJustify",visSectionSmartTag,visRowSmartTag ,visSmartTagYJustify,L"SmartTags.{r}.YJustify",L"0=visSmartTagYJustifyTop;1=visSmartTagYJustifyMiddle;2=visSmartTagYJustifyBottom"},
	{11,L"Action Tags",L"{r}",L"DisplayMode",visSectionSmartTag,visRowSmartTag ,visSmartTagDisplayMode,L"SmartTags.{r}.DisplayMode",L"0=visSmartTagDispModeMouseOver;1=visSmartTagDispModeShapeSelected;2=visSmartTagDispModeAlways"},
	{11,L"Action Tags",L"{r}",L"ButtonFace",visSectionSmartTag,visRowSmartTag ,visSmartTagButtonFace,L"SmartTags.{r}.ButtonFace",L""},
	{11,L"Action Tags",L"{r}",L"Disabled",visSectionSmartTag,visRowSmartTag ,visSmartTagDisabled,L"SmartTags.{r}.Disabled",L"TRUE;FALSE"},
	{11,L"Action Tags",L"{r}",L"Description",visSectionSmartTag,visRowSmartTag ,visSmartTagDescription,L"SmartTags.{r}.Description",L""},
	{0,L"Actions",L"{r}",L"Menu",visSectionAction,visRowAction ,visActionMenu,L"Actions.{r}.Menu",L""},
	{0,L"Actions",L"{r}",L"Prompt",visSectionAction,visRowAction ,visActionPrompt,L"Actions.{r}.Prompt",L""},
	{0,L"Actions",L"{r}",L"Help",visSectionAction,visRowAction ,visActionHelp,L"Actions.{r}.Help",L""},
	{0,L"Actions",L"{r}",L"Action",visSectionAction,visRowAction ,visActionAction,L"Actions.{r}.Action",L""},
	{0,L"Actions",L"{r}",L"Checked",visSectionAction,visRowAction ,visActionChecked,L"Actions.{r}.Checked",L"TRUE;FALSE"},
	{0,L"Actions",L"{r}",L"Disabled",visSectionAction,visRowAction ,visActionDisabled,L"Actions.{r}.Disabled",L"TRUE;FALSE"},
	{11,L"Actions",L"{r}",L"ReadOnly",visSectionAction,visRowAction ,visActionReadOnly,L"Actions.{r}.ReadOnly",L"TRUE;FALSE"},
	{11,L"Actions",L"{r}",L"Invisible",visSectionAction,visRowAction ,visActionInvisible,L"Actions.{r}.Invisible",L"TRUE;FALSE"},
	{11,L"Actions",L"{r}",L"BeginGroup",visSectionAction,visRowAction ,visActionBeginGroup,L"Actions.{r}.BeginGroup",L"TRUE;FALSE"},
	{14,L"Actions",L"{r}",L"FlyoutChild",visSectionAction,visRowAction ,visActionFlyoutChild,L"Actions.{r}.FlyoutChild",L""},
	{11,L"Actions",L"{r}",L"TagName",visSectionAction,visRowAction ,visActionTagName,L"Actions.{r}.TagName",L""},
	{11,L"Actions",L"{r}",L"ButtonFace",visSectionAction,visRowAction ,visActionButtonFace,L"Actions.{r}.ButtonFace",L""},
	{11,L"Actions",L"{r}",L"SortKey",visSectionAction,visRowAction ,visActionSortKey,L"Actions.{r}.SortKey",L""},
	{15,L"Additional Effect Properties",L"",L"ReflectionTrans",visSectionObject,visRowOtherEffectProperties,visReflectionTrans,L"ReflectionTrans",L""},
	{15,L"Additional Effect Properties",L"",L"ReflectionSize",visSectionObject,visRowOtherEffectProperties,visReflectionSize,L"ReflectionSize",L""},
	{15,L"Additional Effect Properties",L"",L"ReflectionDist",visSectionObject,visRowOtherEffectProperties,visReflectionDist,L"ReflectionDist",L""},
	{15,L"Additional Effect Properties",L"",L"ReflectionBlur",visSectionObject,visRowOtherEffectProperties,visReflectionBlur,L"ReflectionBlur",L""},
	{15,L"Additional Effect Properties",L"",L"GlowColor",visSectionObject,visRowOtherEffectProperties,visGlowColor,L"GlowColor",L""},
	{15,L"Additional Effect Properties",L"",L"GlowColorTrans",visSectionObject,visRowOtherEffectProperties,visGlowColorTrans,L"GlowColorTrans",L""},
	{15,L"Additional Effect Properties",L"",L"GlowSize",visSectionObject,visRowOtherEffectProperties,visGlowSize,L"GlowSize",L""},
	{15,L"Additional Effect Properties",L"",L"SoftEdgesSize",visSectionObject,visRowOtherEffectProperties,visSoftEdgesSize,L"SoftEdgesSize",L""},
	{15,L"Additional Effect Properties",L"",L"SketchSeed",visSectionObject,visRowOtherEffectProperties,visSketchSeed,L"SketchSeed",L""},
	{15,L"Additional Effect Properties",L"",L"SketchEnabled",visSectionObject,visRowOtherEffectProperties,visSketchEnabled,L"SketchEnabled",L""},
	{15,L"Additional Effect Properties",L"",L"SketchAmount",visSectionObject,visRowOtherEffectProperties,visSketchAmount,L"SketchAmount",L"0-25"},
	{15,L"Additional Effect Properties",L"",L"SketchLineWeight",visSectionObject,visRowOtherEffectProperties,visSketchLineWeight,L"SketchLineWeight",L""},
	{15,L"Additional Effect Properties",L"",L"SketchLineChange",visSectionObject,visRowOtherEffectProperties,visSketchLineChange,L"SketchLineChange",L""},
	{15,L"Additional Effect Properties",L"",L"SketchFillChange",visSectionObject,visRowOtherEffectProperties,visSketchFillChange,L"SketchFillChange",L""},
	{0,L"Alignment",L"",L"AlignLeft",visSectionObject,visRowAlign,visAlignLeft,L"AlignLeft",L""},
	{0,L"Alignment",L"",L"AlignCenter",visSectionObject,visRowAlign,visAlignCenter,L"AlignCenter",L""},
	{0,L"Alignment",L"",L"AlignRight",visSectionObject,visRowAlign,visAlignRight,L"AlignRight",L""},
	{0,L"Alignment",L"",L"AlignTop",visSectionObject,visRowAlign,visAlignTop,L"AlignTop",L""},
	{0,L"Alignment",L"",L"AlignMiddle",visSectionObject,visRowAlign,visAlignMiddle,L"AlignMiddle",L""},
	{0,L"Alignment",L"",L"AlignBottom",visSectionObject,visRowAlign,visAlignBottom,L"AlignBottom",L""},
	{11,L"Annotation",L"{i}",L"X",visSectionAnnotation,visRowAnnotation ,visAnnotationX,L"Annotation.X[{i}]",L""},
	{11,L"Annotation",L"{i}",L"Y",visSectionAnnotation,visRowAnnotation ,visAnnotationY,L"Annotation.Y[{i}]",L""},
	{11,L"Annotation",L"{i}",L"Date",visSectionAnnotation,visRowAnnotation ,visAnnotationDate,L"Annotation.Date[{i}]",L""},
	{11,L"Annotation",L"{i}",L"Comment",visSectionAnnotation,visRowAnnotation ,visAnnotationComment,L"Annotation.Comment[{i}]",L""},
	{11,L"Annotation",L"{i}",L"LangID",visSectionAnnotation,visRowAnnotation ,visAnnotationLangID,L"Annotation.LangID[{i}]",L""},
	{15,L"Bevel Properties",L"",L"BevelTopType",visSectionObject,visRowBevelProperties,visBevelTopType,L"BevelTopType",L"0-12"},
	{15,L"Bevel Properties",L"",L"BevelTopWidth",visSectionObject,visRowBevelProperties,visBevelTopWidth,L"BevelTopWidth",L""},
	{15,L"Bevel Properties",L"",L"BevelTopHeight",visSectionObject,visRowBevelProperties,visBevelTopHeight,L"BevelTopHeight",L""},
	{15,L"Bevel Properties",L"",L"BevelBottomType",visSectionObject,visRowBevelProperties,visBevelBottomType,L"BevelBottomType",L"0-12"},
	{15,L"Bevel Properties",L"",L"BevelBottomWidth",visSectionObject,visRowBevelProperties,visBevelBottomWidth,L"BevelBottomWidth",L""},
	{15,L"Bevel Properties",L"",L"BevelBottomHeight",visSectionObject,visRowBevelProperties,visBevelBottomHeight,L"BevelBottomHeight",L""},
	{15,L"Bevel Properties",L"",L"BevelDepthColor",visSectionObject,visRowBevelProperties,visBevelDepthColor,L"BevelDepthColor",L""},
	{15,L"Bevel Properties",L"",L"BevelDepthSize",visSectionObject,visRowBevelProperties,visBevelDepthSize,L"BevelDepthSize",L""},
	{15,L"Bevel Properties",L"",L"BevelContourColor",visSectionObject,visRowBevelProperties,visBevelContourColor,L"BevelContourColor",L""},
	{15,L"Bevel Properties",L"",L"BevelContourSize",visSectionObject,visRowBevelProperties,visBevelContourSize,L"BevelContourSize",L""},
	{15,L"Bevel Properties",L"",L"BevelMaterialType",visSectionObject,visRowBevelProperties,visBevelMaterialType,L"BevelMaterialType",L"0-11"},
	{15,L"Bevel Properties",L"",L"BevelLightingType",visSectionObject,visRowBevelProperties,visBevelLightingType,L"BevelLightingType",L"0-15"},
	{15,L"Bevel Properties",L"",L"BevelLightingAngle",visSectionObject,visRowBevelProperties,visBevelLightingAngle,L"BevelLightingAngle",L""},
	{15,L"Change Shape Behavior",L"",L"ReplaceLockShapeData",visSectionObject,visRowReplaceBehaviors,visReplaceLockShapeData,L"ReplaceLockShapeData",L"TRUE;FALSE"},
	{15,L"Change Shape Behavior",L"",L"ReplaceLockText",visSectionObject,visRowReplaceBehaviors,visReplaceLockText,L"ReplaceLockText",L"TRUE;FALSE"},
	{15,L"Change Shape Behavior",L"",L"ReplaceLockFormat",visSectionObject,visRowReplaceBehaviors,visReplaceLockFormat,L"ReplaceLockFormat",L"TRUE;FALSE"},
	{15,L"Change Shape Behavior",L"",L"ReplaceCopyCells",visSectionObject,visRowReplaceBehaviors,visReplaceCopyCells,L"ReplaceCopyCells",L""},
	{0,L"Character",L"{i}",L"Font",visSectionCharacter,visRowCharacter ,visCharacterFont,L"Char.Font[{i}]",L""},
	{0,L"Character",L"{i}",L"Color",visSectionCharacter,visRowCharacter ,visCharacterColor,L"Char.Color[{i}]",L""},
	{0,L"Character",L"{i}",L"Style",visSectionCharacter,visRowCharacter ,visCharacterStyle,L"Char.Style[{i}]",L"Bold=visBold;Italic=visItalic;Underline=visUnderLine;Small caps=visSmallCaps"},
	{0,L"Character",L"{i}",L"Case",visSectionCharacter,visRowCharacter ,visCharacterCase,L"Char.Case[{i}]",L"0=visCaseNormal;1=visCaseAllCaps;2=visCaseInitialCaps"},
	{0,L"Character",L"{i}",L"Pos",visSectionCharacter,visRowCharacter ,visCharacterPos,L"Char.Pos[{i}]",L"0=visPosNormal;1=visPosSuper;2=visPosSub"},
	{0,L"Character",L"{i}",L"FontScale",visSectionCharacter,visRowCharacter ,visCharacterFontScale,L"Char.FontScale[{i}]",L""},
	{0,L"Character",L"{i}",L"Size",visSectionCharacter,visRowCharacter ,visCharacterSize,L"Char.Size[{i}]",L""},
	{0,L"Character",L"{i}",L"DblUnderline",visSectionCharacter,visRowCharacter ,visCharacterDblUnderline,L"Char.DblUnderline[{i}]",L"TRUE;FALSE"},
	{0,L"Character",L"{i}",L"Overline",visSectionCharacter,visRowCharacter ,visCharacterOverline,L"Char.Overline[{i}]",L"TRUE;FALSE"},
	{0,L"Character",L"{i}",L"Strikethru",visSectionCharacter,visRowCharacter ,visCharacterStrikethru,L"Char.Strikethru[{i}]",L"TRUE;FALSE"},
	{11,L"Character",L"{i}",L"DoubleStrikethrough",visSectionCharacter,visRowCharacter ,visCharacterDoubleStrikethrough,L"Char.DoubleStrikethrough[{i}]",L""},
	{12,L"Character",L"{i}",L"RTLText",visSectionCharacter,visRowCharacter ,visCharacterRTLText,L"Char.RTLText[i]",L"TRUE;FALSE"},
	{12,L"Character",L"{i}",L"UseVertical",visSectionCharacter,visRowCharacter ,visCharacterUseVertical,L"Char.UseVertical[i]",L"TRUE;FALSE"},
	{0,L"Character",L"{i}",L"Letterspace",visSectionCharacter,visRowCharacter ,visCharacterLetterspace,L"Char.Letterspace[{i}]",L""},
	{0,L"Character",L"{i}",L"ColorTrans",visSectionCharacter,visRowCharacter ,visCharacterColorTrans,L"Char.ColorTrans[{i}]",L"0-100"},
	{11,L"Character",L"{i}",L"AsianFont",visSectionCharacter,visRowCharacter ,visCharacterAsianFont,L"Char.AsianFont[{i}]",L""},
	{11,L"Character",L"{i}",L"ComplexScriptFont",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptFont,L"Char.ComplexScriptFont[{i}]",L""},
	{11,L"Character",L"{i}",L"LocalizeFont",visSectionCharacter,visRowCharacter ,visCharacterLocalizeFont,L"Char.LocalizeFont[{i}]",L""},
	{11,L"Character",L"{i}",L"ComplexScriptSize",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptSize,L"Char.ComplexScriptSize[{i}]",L""},
	{11,L"Character",L"{i}",L"LangID",visSectionCharacter,visRowCharacter ,visCharacterLangID,L"Char.LangID[{i}]",L""},
	{0,L"Connection Points",L"{i}",L"X",visSectionConnectionPts,visRowConnectionPts ,visX,L"Connections.X{i}",L""},
	{0,L"Connection Points",L"{i}",L"Y",visSectionConnectionPts,visRowConnectionPts ,visY,L"Connections.Y{i}",L""},
	{0,L"Connection Points",L"{i}",L"DirX",visSectionConnectionPts,visRowConnectionPts ,visCnnctA,L"Connections.DirX[{i}]",L""},
	{0,L"Connection Points",L"{i}",L"DirY",visSectionConnectionPts,visRowConnectionPts ,visCnnctB,L"Connections.DirY[{i}]",L""},
	{0,L"Connection Points",L"{i}",L"Type",visSectionConnectionPts,visRowConnectionPts ,visCnnctC,L"Connections.Type[{i}]",L"0=visCnnctTypeInward;1=visCnnctTypeOutward;2=visCnnctTypeInwardOutward"},
	{0,L"Connection Points",L"{i}",L"D",visSectionConnectionPts,visRowConnectionPts ,visCnnctD,L"Connections.D[{i}]",L""},
	{0,L"Controls",L"{r}",L"X",visSectionControls,visRowControl ,visCtlX,L"Controls.{r}.X",L""},
	{0,L"Controls",L"{r}",L"Y",visSectionControls,visRowControl ,visCtlY,L"Controls.{r}.Y",L""},
	{0,L"Controls",L"{r}",L"XDyn",visSectionControls,visRowControl ,visCtlXDyn,L"Controls.{r}.XDyn",L""},
	{0,L"Controls",L"{r}",L"YDyn",visSectionControls,visRowControl ,visCtlYDyn,L"Controls.{r}.YDyn",L""},
	{0,L"Controls",L"{r}",L"XCon",visSectionControls,visRowControl ,visCtlXCon,L"Controls.{r}.XCon",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{0,L"Controls",L"{r}",L"YCon",visSectionControls,visRowControl ,visCtlYCon,L"Controls.{r}.YCon",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{0,L"Controls",L"{r}",L"CanGlue",visSectionControls,visRowControl ,visCtlGlue,L"Controls.{r}.CanGlue",L"TRUE;FALSE"},
	{0,L"Controls",L"{r}",L"Type",visSectionControls,visRowControl ,visCtlType,L"Controls.{r}.Type",L""},
	{0,L"Controls",L"{r}",L"Tip",visSectionControls,visRowControl ,visCtlTip,L"Controls.{r}.Tip",L""},
	{0,L"Document Properties",L"",L"OutputFormat",visSectionObject,visRowDoc,visDocOutputFormat,L"OutputFormat",L"0-2"},
	{0,L"Document Properties",L"",L"LockPreview",visSectionObject,visRowDoc,visDocLockPreview,L"LockPreview",L"TRUE;FALSE"},
	{0,L"Document Properties",L"",L"Metric",visSectionObject,visRowDoc,visDocMetric,L"Metric",L"TRUE;FALSE"},
	{11,L"Document Properties",L"",L"AddMarkup",visSectionObject,visRowDoc,visDocAddMarkup,L"AddMarkup",L"TRUE;FALSE"},
	{11,L"Document Properties",L"",L"ViewMarkup",visSectionObject,visRowDoc,visDocViewMarkup,L"ViewMarkup",L"TRUE;FALSE"},
	{0,L"Document Properties",L"",L"DocLocReplace",visSectionObject,visRowDoc,visDocLockReplace,L"DocLocReplace",L""},
	{0,L"Document Properties",L"",L"NoCoauth",visSectionObject,visRowDoc,visDocNoCoauth,L"NoCoauth",L"TRUE;FALSE"},
	{0,L"Document Properties",L"",L"DocLockDuplicatePage",visSectionObject,visRowDoc,visDocLockDuplicatePage,L"DocLockDuplicatePage",L""},
	{0,L"Document Properties",L"",L"PreviewQuality",visSectionObject,visRowDoc,visDocPreviewQuality,L"PreviewQuality",L"0=visDocPreviewQualityDraft;1=visDocPreviewQualityDetailed"},
	{0,L"Document Properties",L"",L"PreviewScope",visSectionObject,visRowDoc,visDocPreviewScope,L"PreviewScope",L"0=visDocPreviewScope1stPage;1=visDocPreviewScopeNone;2=visDocPreviewScopeAllPages"},
	{11,L"Document Properties",L"",L"DocLangID",visSectionObject,visRowDoc,visDocLangID,L"DocLangID",L""},
	{0,L"Events",L"",L"TheData",visSectionObject,visRowEvent,visEvtCellTheData,L"TheData",L""},
	{0,L"Events",L"",L"TheText",visSectionObject,visRowEvent,visEvtCellTheText,L"TheText",L""},
	{0,L"Events",L"",L"EventDblClick",visSectionObject,visRowEvent,visEvtCellDblClick,L"EventDblClick",L""},
	{0,L"Events",L"",L"EventXFMod",visSectionObject,visRowEvent,visEvtCellXFMod,L"EventXFMod",L""},
	{0,L"Events",L"",L"EventDrop",visSectionObject,visRowEvent,visEvtCellDrop,L"EventDrop",L""},
	{12,L"Events",L"",L"EventMultiDrop",visSectionObject,visRowEvent,visEvtCellMultiDrop,L"EventMultiDrop",L""},
	{0,L"Fill Format",L"",L"FillForegnd",visSectionObject,visRowFill,visFillForegnd,L"FillForegnd",L""},
	{0,L"Fill Format",L"",L"FillBkgnd",visSectionObject,visRowFill,visFillBkgnd,L"FillBkgnd",L""},
	{0,L"Fill Format",L"",L"FillPattern",visSectionObject,visRowFill,visFillPattern,L"FillPattern",L"0-40"},
	{0,L"Fill Format",L"",L"ShdwForegnd",visSectionObject,visRowFill,visFillShdwForegnd,L"ShdwForegnd",L""},
	{0,L"Fill Format",L"",L"ShdwPattern",visSectionObject,visRowFill,visFillShdwPattern,L"ShdwPattern",L"0-40"},
	{0,L"Fill Format",L"",L"FillForegndTrans",visSectionObject,visRowFill,visFillForegndTrans,L"FillForegndTrans",L"0-100"},
	{0,L"Fill Format",L"",L"FillBkgndTrans",visSectionObject,visRowFill,visFillBkgndTrans,L"FillBkgndTrans",L"0-100"},
	{0,L"Fill Format",L"",L"ShdwForegndTrans",visSectionObject,visRowFill,visFillShdwForegndTrans,L"ShdwForegndTrans",L"0-100"},
	{0,L"Fill Format",L"",L"ShdwBkgndTrans",visSectionObject,visRowFill,visFillShdwBkgndTrans,L"ShdwBkgndTrans",L"0-100"},
	{0,L"Fill Format",L"",L"ShapeShdwType",visSectionObject,visRowFill,visFillShdwType,L"ShapeShdwType",L"0=visFSTPageDefault;1=visFSTSimple;2=visFSTOblique"},
	{11,L"Fill Format",L"",L"ShapeShdwOffsetX",visSectionObject,visRowFill,visFillShdwOffsetX,L"ShapeShdwOffsetX",L""},
	{11,L"Fill Format",L"",L"ShapeShdwOffsetY",visSectionObject,visRowFill,visFillShdwOffsetY,L"ShapeShdwOffsetY",L""},
	{11,L"Fill Format",L"",L"ShapeShdwObliqueAngle",visSectionObject,visRowFill,visFillShdwObliqueAngle,L"ShapeShdwObliqueAngle",L""},
	{11,L"Fill Format",L"",L"ShapeShdwScaleFactor",visSectionObject,visRowFill,visFillShdwScaleFactor,L"ShapeShdwScaleFactor",L""},
	{15,L"Fill Format",L"",L"ShapeShdwBlur",visSectionObject,visRowFill,visFillShdwBlur,L"ShapeShdwBlur",L""},
	{15,L"Fill Format",L"",L"ShapeShdwShow",visSectionObject,visRowFill,visFillShdwShow,L"ShapeShdwShow",L"0-2"},
	{0,L"Foreign Image Info",L"",L"ImgOffsetX",visSectionObject,visRowForeign,visFrgnImgOffsetX,L"ImgOffsetX",L""},
	{0,L"Foreign Image Info",L"",L"ImgOffsetY",visSectionObject,visRowForeign,visFrgnImgOffsetY,L"ImgOffsetY",L""},
	{0,L"Foreign Image Info",L"",L"ImgWidth",visSectionObject,visRowForeign,visFrgnImgWidth,L"ImgWidth",L""},
	{0,L"Foreign Image Info",L"",L"ImgHeight",visSectionObject,visRowForeign,visFrgnImgHeight,L"ImgHeight",L""},
	{15,L"Foreign Image Info",L"",L"ClippingPath",visSectionObject,visRowForeign,visFrgnImgClippingPath,L"ClippingPath",L""},
	{0,L"Geometry",L"{i}",L"NoFill",visSectionFirstComponent,visRowComponent,visCompNoFill,L"Geometry{i}.NoFill",L"TRUE;FALSE"},
	{0,L"Geometry",L"{i}",L"NoLine",visSectionFirstComponent,visRowComponent,visCompNoLine,L"Geometry{i}.NoLine",L"TRUE;FALSE"},
	{0,L"Geometry",L"{i}",L"NoShow",visSectionFirstComponent,visRowComponent,visCompNoShow,L"Geometry{i}.NoShow",L"TRUE;FALSE"},
	{0,L"Geometry",L"{i}",L"NoSnap",visSectionFirstComponent,visRowComponent,visCompNoSnap,L"Geometry{i}.NoSnap",L"TRUE;FALSE"},
	{14,L"Geometry",L"{i}",L"NoQuickDrag",visSectionFirstComponent,visRowComponent,visCompNoQuickDrag,L"Geometry{i}.NoQuickDrag",L"TRUE;FALSE"},
	{0,L"Geometry",L"{i}",L"X",visSectionFirstComponent,visRowVertex ,visX,L"Geometry{i}.X{j}",L""},
	{0,L"Geometry",L"{i}",L"Y",visSectionFirstComponent,visRowVertex ,visY,L"Geometry{i}.Y{j}",L""},
	{0,L"Geometry",L"{i}",L"A",visSectionFirstComponent,visRowVertex ,visNURBSKnot,L"Geometry{i}.A{j}",L""},
	{0,L"Geometry",L"{i}",L"B",visSectionFirstComponent,visRowVertex ,visNURBSWeight,L"Geometry{i}.B{j}",L""},
	{0,L"Geometry",L"{i}",L"C",visSectionFirstComponent,visRowVertex ,visNURBSKnotPrev,L"Geometry{i}.C{j}",L""},
	{0,L"Geometry",L"{i}",L"D",visSectionFirstComponent,visRowVertex ,visNURBSWeightPrev,L"Geometry{i}.D{j}",L""},
	{0,L"Geometry",L"{i}",L"E",visSectionFirstComponent,visRowVertex ,visNURBSData,L"Geometry{i}.E{j}",L""},
	{0,L"Glue Info",L"",L"GlueType",visSectionObject,visRowMisc,visGlueType,L"GlueType",L"0=visGlueTypeDefault;2=visGlueTypeWalking;4=visGlueTypeNoWalking;8=visGlueTypeNoWalkingTo"},
	{0,L"Glue Info",L"",L"WalkPreference",visSectionObject,visRowMisc,visWalkPref,L"WalkPreference",L"1=visWalkPrefBegNS;2=visWalkPrefEndNS"},
	{0,L"Glue Info",L"",L"BegTrigger",visSectionObject,visRowGroup,visBegTrigger,L"BegTrigger",L""},
	{0,L"Glue Info",L"",L"EndTrigger",visSectionObject,visRowMisc,visEndTrigger,L"EndTrigger",L""},
	{15,L"Gradient Properties",L"",L"LineGradientDir",visSectionObject,visRowGradientProperties,visLineGradientDir,L"LineGradientDir",L"0-13"},
	{15,L"Gradient Properties",L"",L"LineGradientAngle",visSectionObject,visRowGradientProperties,visLineGradientAngle,L"LineGradientAngle",L""},
	{15,L"Gradient Properties",L"",L"FillGradientDir",visSectionObject,visRowGradientProperties,visFillGradientDir,L"FillGradientDir",L"0-13"},
	{15,L"Gradient Properties",L"",L"FillGradientAngle",visSectionObject,visRowGradientProperties,visFillGradientAngle,L"FillGradientAngle",L""},
	{15,L"Gradient Properties",L"",L"LineGradientEnabled",visSectionObject,visRowGradientProperties,visLineGradientEnabled,L"LineGradientEnabled",L"TRUE;FALSE"},
	{15,L"Gradient Properties",L"",L"FillGradientEnabled",visSectionObject,visRowGradientProperties,visFillGradientEnabled,L"FillGradientEnabled",L"TRUE;FALSE"},
	{15,L"Gradient Properties",L"",L"RotateGradientWithShape",visSectionObject,visRowGradientProperties,visRotateGradientWithShape,L"RotateGradientWithShape",L"TRUE;FALSE"},
	{15,L"Gradient Properties",L"",L"UseGroupGradient",visSectionObject,visRowGradientProperties,visUseGroupGradient,L"UseGroupGradient",L""},
	{0,L"Group Properties",L"",L"SelectMode",visSectionObject,visRowGroup,visGroupSelectMode,L"SelectMode",L"0=visGrpSelModeGroupOnly;1=visGrpSelModeGroup1st;2=visGrpSelModeMembers1st"},
	{0,L"Group Properties",L"",L"DisplayMode",visSectionObject,visRowGroup,visGroupDisplayMode,L"DisplayMode",L"0=visGrpDispModeNone;1=visGrpDispModeBack;2=visGrpDispModeFront"},
	{0,L"Group Properties",L"",L"IsDropTarget",visSectionObject,visRowGroup,visGroupIsDropTarget,L"IsDropTarget",L"TRUE;FALSE"},
	{0,L"Group Properties",L"",L"IsSnapTarget",visSectionObject,visRowGroup,visGroupIsSnapTarget,L"IsSnapTarget",L"TRUE;FALSE"},
	{0,L"Group Properties",L"",L"IsTextEditTarget",visSectionObject,visRowGroup,visGroupIsTextEditTarget,L"IsTextEditTarget",L"TRUE;FALSE"},
	{0,L"Group Properties",L"",L"DontMoveChildren",visSectionObject,visRowGroup,visGroupDontMoveChildren,L"DontMoveChildren",L"TRUE;FALSE"},
	{0,L"Hyperlinks",L"{r}",L"Description",visSectionHyperlink,visRow1stHyperlink ,visHLinkDescription,L"Hyperlink.{r}.Description",L""},
	{0,L"Hyperlinks",L"{r}",L"Address",visSectionHyperlink,visRow1stHyperlink ,visHLinkAddress,L"Hyperlink.{r}.Address",L""},
	{0,L"Hyperlinks",L"{r}",L"SubAddress",visSectionHyperlink,visRow1stHyperlink ,visHLinkSubAddress,L"Hyperlink.{r}.SubAddress",L""},
	{0,L"Hyperlinks",L"{r}",L"ExtraInfo",visSectionHyperlink,visRow1stHyperlink ,visHLinkExtraInfo,L"Hyperlink.{r}.ExtraInfo",L""},
	{0,L"Hyperlinks",L"{r}",L"Frame",visSectionHyperlink,visRow1stHyperlink ,visHLinkFrame,L"Hyperlink.{r}.Frame",L""},
	{0,L"Hyperlinks",L"{r}",L"NewWindow",visSectionHyperlink,visRow1stHyperlink ,visHLinkNewWin,L"Hyperlink.{r}.NewWindow",L"TRUE;FALSE"},
	{0,L"Hyperlinks",L"{r}",L"Default",visSectionHyperlink,visRow1stHyperlink ,visHLinkDefault,L"Hyperlink.{r}.Default",L""},
	{11,L"Hyperlinks",L"{r}",L"Invisible",visSectionHyperlink,visRow1stHyperlink ,visHLinkInvisible,L"Hyperlink.{r}.Invisible",L"TRUE;FALSE"},
	{11,L"Hyperlinks",L"{r}",L"SortKey",visSectionHyperlink,visRow1stHyperlink ,visHLinkSortKey,L"Hyperlink.{r}.SortKey",L""},
	{0,L"Image Properties",L"",L"Gamma",visSectionObject,visRowImage,visImageGamma,L"Gamma",L""},
	{0,L"Image Properties",L"",L"Contrast",visSectionObject,visRowImage,visImageContrast,L"Contrast",L""},
	{0,L"Image Properties",L"",L"Brightness",visSectionObject,visRowImage,visImageBrightness,L"Brightness",L""},
	{0,L"Image Properties",L"",L"Sharpen",visSectionObject,visRowImage,visImageSharpen,L"Sharpen",L""},
	{0,L"Image Properties",L"",L"Blur",visSectionObject,visRowImage,visImageBlur,L"Blur",L""},
	{0,L"Image Properties",L"",L"Denoise",visSectionObject,visRowImage,visImageDenoise,L"Denoise",L""},
	{0,L"Image Properties",L"",L"Transparency",visSectionObject,visRowImage,visImageTransparency,L"Transparency",L"0-100"},
	{0,L"Layers",L"{i}",L"Name",visSectionLayer,visRowLayer ,visLayerName,L"Layers.Name[{i}]",L""},
	{0,L"Layers",L"{i}",L"Color",visSectionLayer,visRowLayer ,visLayerColor,L"Layers.Color[{i}]",L""},
	{0,L"Layers",L"{i}",L"Visible",visSectionLayer,visRowLayer ,visLayerVisible,L"Layers.Visible[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"Active",visSectionLayer,visRowLayer ,visLayerActive,L"Layers.Active[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"Locked",visSectionLayer,visRowLayer ,visLayerLock,L"Layers.Locked[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"Snap",visSectionLayer,visRowLayer ,visLayerSnap,L"Layers.Snap[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"Glue",visSectionLayer,visRowLayer ,visLayerGlue,L"Layers.Glue[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"Print",visSectionLayer,visRowLayer ,visDocPreviewScope,L"Layers.Print[{i}]",L"TRUE;FALSE"},
	{0,L"Layers",L"{i}",L"ColorTrans",visSectionLayer,visRowLayer ,visLayerColorTrans,L"Layers.ColorTrans[{i}]",L"0-100"},
	{0,L"Line Format",L"",L"LineWeight",visSectionObject,visRowLine,visLineWeight,L"LineWeight",L""},
	{0,L"Line Format",L"",L"LineColor",visSectionObject,visRowLine,visLineColor,L"LineColor",L""},
	{0,L"Line Format",L"",L"LinePattern",visSectionObject,visRowLine,visLinePattern,L"LinePattern",L"0-23"},
	{0,L"Line Format",L"",L"Rounding",visSectionObject,visRowLine,visLineRounding,L"Rounding",L""},
	{0,L"Line Format",L"",L"EndArrowSize",visSectionObject,visRowLine,visLineEndArrowSize,L"EndArrowSize",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{0,L"Line Format",L"",L"BeginArrow",visSectionObject,visRowLine,visLineBeginArrow,L"BeginArrow",L"0-45"},
	{0,L"Line Format",L"",L"EndArrow",visSectionObject,visRowLine,visLineEndArrow,L"EndArrow",L"0-45"},
	{0,L"Line Format",L"",L"LineCap",visSectionObject,visRowLine,visLineEndCap,L"LineCap",L"0-2"},
	{0,L"Line Format",L"",L"BeginArrowSize",visSectionObject,visRowLine,visLineBeginArrowSize,L"BeginArrowSize",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{0,L"Line Format",L"",L"LineColorTrans",visSectionObject,visRowLine,visLineColorTrans,L"LineColorTrans",L"0-100"},
	{15,L"Line Format",L"",L"CompoundType",visSectionObject,visRowLine,visCompoundType,L"CompoundType",L"0-4"},
	{0,L"Miscellaneous",L"",L"NoObjHandles",visSectionObject,visRowMisc,visNoObjHandles,L"NoObjHandles",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"NonPrinting",visSectionObject,visRowMisc,visNonPrinting,L"NonPrinting",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"NoCtlHandles",visSectionObject,visRowMisc,visNoCtlHandles,L"NoCtlHandles",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"NoAlignBox",visSectionObject,visRowMisc,visNoAlignBox,L"NoAlignBox",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"UpdateAlignBox",visSectionObject,visRowMisc,visUpdateAlignBox,L"UpdateAlignBox",L""},
	{0,L"Miscellaneous",L"",L"HideText",visSectionObject,visRowMisc,visHideText,L"HideText",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"DynFeedback",visSectionObject,visRowMisc,visDynFeedback,L"DynFeedback",L"0=visDynFBDefault;1=visDynFBUCon3Leg;2=visDynFBUCon5Leg"},
	{0,L"Miscellaneous",L"",L"ObjType",visSectionObject,visRowMisc,visLOFlags,L"ObjType",L"0=visLOFlagsVisDecides;1=visLOFlagsPlacable;2=visLOFlagsRoutable;4=visLOFlagsDont;8=visLOFlagsPNRGroup"},
	{0,L"Miscellaneous",L"",L"Comment",visSectionObject,visRowMisc,visComment,L"Comment",L""},
	{0,L"Miscellaneous",L"",L"IsDropSource",visSectionObject,visRowMisc,visDropSource,L"IsDropSource",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"NoLiveDynamics",visSectionObject,visRowMisc,visNoLiveDynamics,L"NoLiveDynamics",L"TRUE;FALSE"},
	{11,L"Miscellaneous",L"",L"LocalizeMerge",visSectionObject,visRowMisc,visObjLocalizeMerge,L"LocalizeMerge",L"TRUE;FALSE"},
	{0,L"Miscellaneous",L"",L"NoProofing",visSectionObject,visRowMisc,visObjNoProofing,L"NoProofing",L"TRUE;FALSE"},
	{11,L"Miscellaneous",L"",L"Calendar",visSectionObject,visRowMisc,visObjCalendar,L"Calendar",L"0=Western;1=Arabic Hijri;2=Hebrew Lunar;3=Taiwan Calendar;4=Japanese Emperor Reign;5=Thai Buddhist;6=Korean Danki;7=Saka Era;8=English transliterated;9=French transliterated"},
	{11,L"Miscellaneous",L"",L"LangID",visSectionObject,visRowMisc,visObjLangID,L"LangID",L""},
	{11,L"Miscellaneous",L"",L"DropOnPageScale",visSectionObject,visRowMisc,visObjDropOnPageScale,L"DropOnPageScale",L""},
	{0,L"Page Layout",L"",L"ResizePage",visSectionObject,visRowPageLayout,visPLOResizePage,L"ResizePage",L"TRUE;FALSE"},
	{0,L"Page Layout",L"",L"EnableGrid",visSectionObject,visRowPageLayout,visPLOEnableGrid,L"EnableGrid",L"TRUE;FALSE"},
	{0,L"Page Layout",L"",L"DynamicsOff",visSectionObject,visRowPageLayout,visPLODynamicsOff,L"DynamicsOff",L"TRUE;FALSE"},
	{0,L"Page Layout",L"",L"CtrlAsInput",visSectionObject,visRowPageLayout,visPLOCtrlAsInput,L"CtrlAsInput",L"TRUE;FALSE"},
	{0,L"Page Layout",L"",L"AvoidPageBreaks",visSectionObject,visRowPageLayout,visPLOAvoidPageBreaks,L"AvoidPageBreaks",L""},
	{0,L"Page Layout",L"",L"PlaceStyle",visSectionObject,visRowPageLayout,visPLOPlaceStyle,L"PlaceStyle",L""},
	{0,L"Page Layout",L"",L"RouteStyle",visSectionObject,visRowPageLayout,visPLORouteStyle,L"RouteStyle",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{0,L"Page Layout",L"",L"PlaceDepth",visSectionObject,visRowPageLayout,visPLOPlaceDepth,L"PlaceDepth",L"0=visPLOPlaceDepthDefault;1=visPLOPlaceDepthMedium;2=visPLOPlaceDepthDeep;3=visPLOPlaceDepthShallow"},
	{0,L"Page Layout",L"",L"PlowCode",visSectionObject,visRowPageLayout,visPLOPlowCode,L"PlowCode",L"0=visPLOPlowNone;1=visPLOPlowAll"},
	{0,L"Page Layout",L"",L"LineJumpCode",visSectionObject,visRowPageLayout,visPLOJumpCode,L"LineJumpCode",L"0=visPLOJumpNone;1=visPLOJumpHorizontal;2=visPLOJumpVertical;3=visPLOJumpLastRouted;4=visPLOJumpDisplayOrder;5=visPLOJumpReverseDisplayOrder"},
	{0,L"Page Layout",L"",L"LineJumpStyle",visSectionObject,visRowPageLayout,visPLOJumpStyle,L"LineJumpStyle",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{0,L"Page Layout",L"",L"PageLineJumpDirX",visSectionObject,visRowPageLayout,visPLOJumpDirX,L"PageLineJumpDirX",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{0,L"Page Layout",L"",L"PageLineJumpDirY",visSectionObject,visRowPageLayout,visPLOJumpDirY,L"PageLineJumpDirY",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{0,L"Page Layout",L"",L"LineToNodeX",visSectionObject,visRowPageLayout,visPLOLineToNodeX,L"LineToNodeX",L""},
	{0,L"Page Layout",L"",L"LineToNodeY",visSectionObject,visRowPageLayout,visPLOLineToNodeY,L"LineToNodeY",L""},
	{0,L"Page Layout",L"",L"BlockSizeX",visSectionObject,visRowPageLayout,visPLOBlockSizeX,L"BlockSizeX",L""},
	{0,L"Page Layout",L"",L"BlockSizeY",visSectionObject,visRowPageLayout,visPLOBlockSizeY,L"BlockSizeY",L""},
	{0,L"Page Layout",L"",L"AvenueSizeX",visSectionObject,visRowPageLayout,visPLOAvenueSizeX,L"AvenueSizeX",L""},
	{0,L"Page Layout",L"",L"AvenueSizeY",visSectionObject,visRowPageLayout,visPLOAvenueSizeY,L"AvenueSizeY",L""},
	{0,L"Page Layout",L"",L"LineToLineX",visSectionObject,visRowPageLayout,visPLOLineToLineX,L"LineToLineX",L""},
	{0,L"Page Layout",L"",L"LineToLineY",visSectionObject,visRowPageLayout,visPLOLineToLineY,L"LineToLineY",L""},
	{0,L"Page Layout",L"",L"LineJumpFactorX",visSectionObject,visRowPageLayout,visPLOJumpFactorX,L"LineJumpFactorX",L""},
	{0,L"Page Layout",L"",L"LineJumpFactorY",visSectionObject,visRowPageLayout,visPLOJumpFactorY,L"LineJumpFactorY",L""},
	{0,L"Page Layout",L"",L"LineAdjustFrom",visSectionObject,visRowPageLayout,visPLOLineAdjustFrom,L"LineAdjustFrom",L"0=visPLOLineAdjustFromNotRelated;1=visPLOLineAdjustFromAll;2=visPLOLineAdjustFromNone;3=visPLOLineAdjustFromRoutingDefault"},
	{0,L"Page Layout",L"",L"LineAdjustTo",visSectionObject,visRowPageLayout,visPLOLineAdjustTo,L"LineAdjustTo",L"0=visPLOLineAdjustToDefault;1=visPLOLineAdjustToAll;2=visPLOLineAdjustToNone;3=visPLOLineAdjustToRelated"},
	{0,L"Page Layout",L"",L"PlaceFlip",visSectionObject,visRowPageLayout,visPLOPlaceFlip,L"PlaceFlip",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{0,L"Page Layout",L"",L"LineRouteExt",visSectionObject,visRowPageLayout,visPLOLineRouteExt,L"LineRouteExt",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{0,L"Page Layout",L"",L"PageShapeSplit",visSectionObject,visRowPageLayout,visPLOSplit,L"PageShapeSplit",L"0=visPLOSplitNone;1=visPLOSplitAllow"},
	{0,L"Page Properties",L"",L"PageWidth",visSectionObject,visRowPage,visPageWidth,L"PageWidth",L""},
	{0,L"Page Properties",L"",L"PageHeight",visSectionObject,visRowPage,visPageHeight,L"PageHeight",L""},
	{0,L"Page Properties",L"",L"ShdwOffsetX",visSectionObject,visRowPage,visPageShdwOffsetX,L"ShdwOffsetX",L""},
	{0,L"Page Properties",L"",L"ShdwOffsetY",visSectionObject,visRowPage,visPageShdwOffsetY,L"ShdwOffsetY",L""},
	{0,L"Page Properties",L"",L"PageScale",visSectionObject,visRowPage,visPageScale,L"PageScale",L""},
	{0,L"Page Properties",L"",L"DrawingScale",visSectionObject,visRowPage,visPageDrawingScale,L"DrawingScale",L""},
	{0,L"Page Properties",L"",L"DrawingSizeType",visSectionObject,visRowPage,visPageDrawSizeType,L"DrawingSizeType",L"0=visPrintSetup;1=visTight;2=visStandard;3=visCustom;4=visLogical;5=visDSMetric;6=visDSEngr;7=visDSArch"},
	{0,L"Page Properties",L"",L"DrawingScaleType",visSectionObject,visRowPage,visPageDrawScaleType,L"DrawingScaleType",L"0=visNoScale;1=visArchitectural;2=visEngineering;3=visScaleCustom;4=visScaleMetric;5=visScaleMechanical"},
	{0,L"Page Properties",L"",L"InhibitSnap",visSectionObject,visRowPage,visPageInhibitSnap,L"InhibitSnap",L"TRUE;FALSE"},
	{0,L"Page Properties",L"",L"PageLockReplace",visSectionObject,visRowPage,visPageLockReplace,L"PageLockReplace",L"TRUE;FALSE"},
	{0,L"Page Properties",L"",L"PageLockDuplicate",visSectionObject,visRowPage,visPageLockDuplicate,L"PageLockDuplicate",L""},
	{0,L"Page Properties",L"",L"UIVisibility",visSectionObject,visRowPage,visPageUIVisibility,L"UIVisibility",L"0=visUIVNormal;1=visUIVHidden"},
	{0,L"Page Properties",L"",L"ShdwType",visSectionObject,visRowPage,visPageShdwType,L"ShdwType",L"1=visFSTSimple;2=visFSTOblique;3=visFSTInner"},
	{0,L"Page Properties",L"",L"ShdwObliqueAngle",visSectionObject,visRowPage,visPageShdwObliqueAngle,L"ShdwObliqueAngle",L""},
	{0,L"Page Properties",L"",L"ShdwScaleFactor",visSectionObject,visRowPage,visPageShdwScaleFactor,L"ShdwScaleFactor",L""},
	{14,L"Page Properties",L"",L"DrawingResizeType",visSectionObject,visRowPage,visPageDrawResizeType,L"DrawingResizeType",L""},
	{0,L"Paragraph",L"{i}",L"IndFirst",visSectionParagraph,visRowParagraph ,visIndentFirst,L"Para.IndFirst[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"IndLeft",visSectionParagraph,visRowParagraph ,visIndentLeft,L"Para.IndLeft[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"IndRight",visSectionParagraph,visRowParagraph ,visIndentRight,L"Para.IndRight[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"SpLine",visSectionParagraph,visRowParagraph ,visSpaceLine,L"Para.SpLine[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"SpBefore",visSectionParagraph,visRowParagraph ,visSpaceBefore,L"Para.SpBefore[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"SpAfter",visSectionParagraph,visRowParagraph ,visSpaceAfter,L"Para.SpAfter[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"HorzAlign",visSectionParagraph,visRowParagraph ,visHorzAlign,L"Para.HorzAlign[{i}]",L"0=visHorzLeft;1=visHorzCenter;2=visHorzRight;3=visHorzJustify;4=visHorzForce"},
	{0,L"Paragraph",L"{i}",L"Bullet",visSectionParagraph,visRowParagraph ,visBulletIndex,L"Para.Bullet[{i}]",L"0-7"},
	{0,L"Paragraph",L"{i}",L"BulletStr",visSectionParagraph,visRowParagraph ,visBulletString,L"Para.BulletStr[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"BulletFont",visSectionParagraph,visRowParagraph ,visBulletFont,L"Para.BulletFont[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"BulletFontSize",visSectionParagraph,visRowParagraph ,visBulletFontSize,L"Para.BulletFontSize[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"TextPosAfterBullet",visSectionParagraph,visRowParagraph ,visTextPosAfterBullet,L"Para.TextPosAfterBullet[{i}]",L""},
	{0,L"Paragraph",L"{i}",L"Flags",visSectionParagraph,visRowParagraph ,visFlags,L"Para.Flags[{i}]",L"0;1"},
	{0,L"Print Properties",L"",L"PageLeftMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesLeftMargin,L"PageLeftMargin",L""},
	{0,L"Print Properties",L"",L"PageRightMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesRightMargin,L"PageRightMargin",L""},
	{0,L"Print Properties",L"",L"PageTopMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesTopMargin,L"PageTopMargin",L""},
	{0,L"Print Properties",L"",L"PageBottomMargin",visSectionObject,visRowPrintProperties,visPrintPropertiesBottomMargin,L"PageBottomMargin",L""},
	{0,L"Print Properties",L"",L"ScaleX",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleX,L"ScaleX",L""},
	{0,L"Print Properties",L"",L"ScaleY",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleY,L"ScaleY",L""},
	{0,L"Print Properties",L"",L"PagesX",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesX,L"PagesX",L""},
	{0,L"Print Properties",L"",L"PagesY",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesY,L"PagesY",L""},
	{0,L"Print Properties",L"",L"CenterX",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterX,L"CenterX",L"TRUE;FALSE"},
	{0,L"Print Properties",L"",L"CenterY",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterY,L"CenterY",L"TRUE;FALSE"},
	{0,L"Print Properties",L"",L"OnPage",visSectionObject,visRowPrintProperties,visPrintPropertiesOnPage,L"OnPage",L"TRUE;FALSE"},
	{0,L"Print Properties",L"",L"PrintGrid",visSectionObject,visRowPrintProperties,visPrintPropertiesPrintGrid,L"PrintGrid",L"TRUE;FALSE"},
	{0,L"Print Properties",L"",L"PrintPageOrientation",visSectionObject,visRowPrintProperties,visPrintPropertiesPageOrientation,L"PrintPageOrientation",L"0=visPPOSameAsPrinter;1=visPPOPortrait;2=visPPOLandscape"},
	{0,L"Print Properties",L"",L"PaperKind",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperKind,L"PaperKind",L""},
	{0,L"PrintProperties",L"",L"PaperSource",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperSource,L"PaperSource",L""},
	{0,L"Protection",L"",L"LockWidth",visSectionObject,visRowLock,visLockWidth,L"LockWidth",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockHeight",visSectionObject,visRowLock,visLockHeight,L"LockHeight",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockMoveX",visSectionObject,visRowLock,visLockMoveX,L"LockMoveX",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockMoveY",visSectionObject,visRowLock,visLockMoveY,L"LockMoveY",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockAspect",visSectionObject,visRowLock,visLockAspect,L"LockAspect",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockDelete",visSectionObject,visRowLock,visLockDelete,L"LockDelete",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockBegin",visSectionObject,visRowLock,visLockBegin,L"LockBegin",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockEnd",visSectionObject,visRowLock,visLockEnd,L"LockEnd",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockRotate",visSectionObject,visRowLock,visLockRotate,L"LockRotate",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockCrop",visSectionObject,visRowLock,visLockCrop,L"LockCrop",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockVtxEdit",visSectionObject,visRowLock,visLockVtxEdit,L"LockVtxEdit",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockTextEdit",visSectionObject,visRowLock,visLockTextEdit,L"LockTextEdit",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockFormat",visSectionObject,visRowLock,visLockFormat,L"LockFormat",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockGroup",visSectionObject,visRowLock,visLockGroup,L"LockGroup",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockCalcWH",visSectionObject,visRowLock,visLockCalcWH,L"LockCalcWH",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockSelect",visSectionObject,visRowLock,visLockSelect,L"LockSelect",L"TRUE;FALSE"},
	{0,L"Protection",L"",L"LockCustProp",visSectionObject,visRowLock,visLockCustProp,L"LockCustProp",L"TRUE;FALSE"},
	{12,L"Protection",L"",L"LockFromGroupFormat",visSectionObject,visRowLock,visLockFromGroupFormat,L"LockFromGroupFormat",L""},
	{12,L"Protection",L"",L"LockThemeColors",visSectionObject,visRowLock,visLockThemeColors,L"LockThemeColors",L""},
	{12,L"Protection",L"",L"LockThemeEffects",visSectionObject,visRowLock,visLockThemeEffects,L"LockThemeEffects",L""},
	{15,L"Protection",L"",L"LockThemeConnectors",visSectionObject,visRowLock,visLockThemeConnectors,L"LockThemeConnectors",L"TRUE;FALSE"},
	{15,L"Protection",L"",L"LockThemeFonts",visSectionObject,visRowLock,visLockThemeFonts,L"LockThemeFonts",L"TRUE;FALSE"},
	{15,L"Protection",L"",L"LockThemeIndex",visSectionObject,visRowLock,visLockThemeIndex,L"LockThemeIndex",L"TRUE;FALSE"},
	{15,L"Protection",L"",L"LockReplace",visSectionObject,visRowLock,visLockReplace,L"LockReplace",L"TRUE;FALSE"},
	{15,L"Protection",L"",L"LockVariation",visSectionObject,visRowLock,visLockVariation,L"LockVariation",L""},
	{15,L"Quick Style",L"",L"QuickStyleLineColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineColor,L"QuickStyleLineColor",L"0-7"},
	{15,L"Quick Style",L"",L"QuickStyleFillColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillColor,L"QuickStyleFillColor",L"0-7"},
	{15,L"Quick Style",L"",L"QuickStyleShadowColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleShadowColor,L"QuickStyleShadowColor",L"0-7"},
	{15,L"Quick Style",L"",L"QuickStyleFontColor",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontColor,L"QuickStyleFontColor",L"0;1"},
	{15,L"Quick Style",L"",L"QuickStyleLineMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineMatrix,L"QuickStyleLineMatrix",L""},
	{15,L"Quick Style",L"",L"QuickStyleFillMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillMatrix,L"QuickStyleFillMatrix",L""},
	{15,L"Quick Style",L"",L"QuickStyleEffectsMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleEffectsMatrix,L"QuickStyleEffectsMatrix",L""},
	{15,L"Quick Style",L"",L"QuickStyleFontMatrix",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontMatrix,L"QuickStyleFontMatrix",L""},
	{15,L"Quick Style",L"",L"QuickStyleType",visSectionObject,visRowQuickStyleProperties,visQuickStyleType,L"QuickStyleType",L"0-3"},
	{15,L"Quick Style",L"",L"QuickStyleVariation",visSectionObject,visRowQuickStyleProperties,visQuickStyleVariation,L"QuickStyleVariation",L"0;1;2;4;8"},
	{0,L"Reviewer",L"{i}",L"Name",visSectionReviewer,visRowReviewer ,visReviewerName,L"Reviewer.Name[{i}]",L""},
	{0,L"Reviewer",L"{i}",L"Initials",visSectionReviewer,visRowReviewer ,visReviewerInitials,L"Reviewer.Initials[{i}]",L""},
	{0,L"Reviewer",L"{i}",L"Color",visSectionReviewer,visRowReviewer ,visReviewerColor,L"Reviewer.Color[{i}]",L""},
	{0,L"Ruler & Grid",L"",L"XRulerDensity",visSectionObject,visRowRulerGrid,visXRulerDensity,L"XRulerDensity",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{0,L"Ruler & Grid",L"",L"YRulerDensity",visSectionObject,visRowRulerGrid,visYRulerDensity,L"YRulerDensity",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{0,L"Ruler & Grid",L"",L"XRulerOrigin",visSectionObject,visRowRulerGrid,visXRulerOrigin,L"XRulerOrigin",L""},
	{0,L"Ruler & Grid",L"",L"YRulerOrigin",visSectionObject,visRowRulerGrid,visYRulerOrigin,L"YRulerOrigin",L""},
	{0,L"Ruler & Grid",L"",L"XGridDensity",visSectionObject,visRowRulerGrid,visXGridDensity,L"XGridDensity",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{0,L"Ruler & Grid",L"",L"YGridDensity",visSectionObject,visRowRulerGrid,visYGridDensity,L"YGridDensity",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{0,L"Ruler & Grid",L"",L"XGridSpacing",visSectionObject,visRowRulerGrid,visXGridSpacing,L"XGridSpacing",L""},
	{0,L"Ruler & Grid",L"",L"YGridSpacing",visSectionObject,visRowRulerGrid,visYGridSpacing,L"YGridSpacing",L""},
	{0,L"Ruler & Grid",L"",L"XGridOrigin",visSectionObject,visRowRulerGrid,visXGridOrigin,L"XGridOrigin",L""},
	{0,L"Ruler & Grid",L"",L"YGridOrigin",visSectionObject,visRowRulerGrid,visYGridOrigin,L"YGridOrigin",L""},
	{0,L"Scratch",L"{i}",L"X",visSectionScratch,visRowScratch,visScratchX,L"Scratch.X{i}",L""},
	{0,L"Scratch",L"{i}",L"Y",visSectionScratch,visRowScratch,visScratchY,L"Scratch.Y{i}",L""},
	{0,L"Scratch",L"{i}",L"A",visSectionScratch,visRowScratch,visScratchA,L"Scratch.A{i}",L""},
	{0,L"Scratch",L"{i}",L"B",visSectionScratch,visRowScratch,visScratchB,L"Scratch.B{i}",L""},
	{0,L"Scratch",L"{i}",L"C",visSectionScratch,visRowScratch,visScratchC,L"Scratch.C{i}",L""},
	{0,L"Scratch",L"{i}",L"D",visSectionScratch,visRowScratch,visScratchD,L"Scratch.D{i}",L""},
	{0,L"Shape Data",L"{r}",L"Value",visSectionProp,visRowProp ,visCustPropsValue,L"Prop.{r}.Value",L""},
	{0,L"Shape Data",L"{r}",L"Prompt",visSectionProp,visRowProp ,visCustPropsPrompt,L"Prop.{r}.Prompt",L""},
	{0,L"Shape Data",L"{r}",L"Label",visSectionProp,visRowProp ,visCustPropsLabel,L"Prop.{r}.Label",L""},
	{0,L"Shape Data",L"{r}",L"Format",visSectionProp,visRowProp ,visCustPropsFormat,L"Prop.{r}.Format",L""},
	{0,L"Shape Data",L"{r}",L"SortKey",visSectionProp,visRowProp ,visCustPropsSortKey,L"Prop.{r}.SortKey",L""},
	{0,L"Shape Data",L"{r}",L"Type",visSectionProp,visRowProp ,visCustPropsType,L"Prop.{r}.Type",L"0=visPropTypeString;1=visPropTypeListFix;2=visPropTypeNumber;3=visPropTypeBool;4=visPropTypeListVar;5=visPropTypeDate;6=visPropTypeDuration;7=visPropTypeCurrency"},
	{0,L"Shape Data",L"{r}",L"Invisible",visSectionProp,visRowProp ,visCustPropsInvis,L"Prop.{r}.Invisible",L"TRUE;FALSE"},
	{0,L"Shape Data",L"{r}",L"Ask",visSectionProp,visRowProp ,visCustPropsAsk,L"Prop.{r}.Ask",L"TRUE;FALSE"},
	{11,L"Shape Data",L"{r}",L"LangID",visSectionProp,visRowProp ,visCustPropsLangID,L"Prop.{r}.LangID",L""},
	{11,L"Shape Data",L"{r}",L"Calendar",visSectionProp,visRowProp ,visCustPropsCalendar,L"Prop.{r}.Calendar",L"0=Western;1=Arabic Hijri;2=Hebrew Lunar;3=Taiwan Calendar;4=Japanese Emperor Reign;5=Thai Buddhist;6=Korean Danki;7=Saka Era;8=English transliterated;9=French transliterated"},
	{0,L"Shape Layout",L"",L"ShapePermeableX",visSectionObject,visRowShapeLayout,visSLOPermX,L"ShapePermeableX",L"TRUE;FALSE"},
	{0,L"Shape Layout",L"",L"ShapePermeableY",visSectionObject,visRowShapeLayout,visSLOPermY,L"ShapePermeableY",L"TRUE;FALSE"},
	{0,L"Shape Layout",L"",L"ShapePermeablePlace",visSectionObject,visRowShapeLayout,visSLOPermeablePlace,L"ShapePermeablePlace",L"TRUE;FALSE"},
	{14,L"Shape Layout",L"",L"Relationships",visSectionObject,visRowShapeLayout,visSLORelationships,L"Relationships",L""},
	{0,L"Shape Layout",L"",L"ShapeFixedCode",visSectionObject,visRowShapeLayout,visSLOFixedCode,L"ShapeFixedCode",L"1=visSLOFixedPlacement;2=visSLOFixedPlow;4=visSLOFixedPermeablePlow;32=visSLOFixedConnPtsIgnore;64=visSLOFixedConnPtsOnly;128=visSLOFixedNoFoldToShape"},
	{0,L"Shape Layout",L"",L"ShapePlowCode",visSectionObject,visRowShapeLayout,visSLOPlowCode,L"ShapePlowCode",L"0=visSLOPlowDefault;1=visSLOPlowNever;2=visSLOPlowAlways"},
	{0,L"Shape Layout",L"",L"ShapeRouteStyle",visSectionObject,visRowShapeLayout,visSLORouteStyle,L"ShapeRouteStyle",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{12,L"Shape Layout",L"",L"ShapePlaceStyle",visSectionObject,visRowShapeLayout,visSLOPlaceStyle,L"ShapePlaceStyle",L"visLOPlaceBottomToTop;visLOPlaceCircular;visLOPlaceCompactDownLeft;visLOPlaceCompactDownRight;visLOPlaceCompactLeftDown;visLOPlaceCompactLeftUp;visLOPlaceCompactRightDown;visLOPlaceCompactRightUp;visLOPlaceCompactUpLeft;visLOPlaceCompactUpRight;visLOPlaceDefault;visLOPlaceHierarchyBottomToTopCenter;visLOPlaceHierarchyBottomToTopLeft;visLOPlaceHierarchyBottomToTopRight;visLOPlaceHierarchyLeftToRightBottom;visLOPlaceHierarchyLeftToRightMiddle;visLOPlaceHierarchyLeftToRightTop;visLOPlaceHierarchyRightToLeftBottom;visLOPlaceHierarchyRightToLeftMiddle;visLOPlaceHierarchyRightToLeftTop;visLOPlaceHierarchyTopToBottomCenter;visLOPlaceHierarchyTopToBottomLeft;visLOPlaceHierarchyTopToBottomRight;visLOPlaceLeftToRight;visLOPlaceParentDefault;visLOPlaceRadial;visLOPlaceRightToLeft;visLOPlaceTopToBottom"},
	{0,L"Shape Layout",L"",L"ConFixedCode",visSectionObject,visRowShapeLayout,visSLOConFixedCode,L"ConFixedCode",L"0=visSLOConFixedRerouteFreely;1=visSLOConFixedRerouteAsNeeded;2=visSLOConFixedRerouteNever;3=visSLOConFixedRerouteOnCrossover;4=visSLOConFixedByAlgFrom;5=visSLOConFixedByAlgTo;6=visSLOConFixedByAlgFromTo"},
	{0,L"Shape Layout",L"",L"ConLineJumpCode",visSectionObject,visRowShapeLayout,visSLOJumpCode,L"ConLineJumpCode",L"0=visSLOJumpDefault;1=visSLOJumpNever;2=visSLOJumpAlways;3=visSLOJumpOther;4=visSLOJumpNeither"},
	{0,L"Shape Layout",L"",L"ConLineJumpStyle",visSectionObject,visRowShapeLayout,visSLOJumpStyle,L"ConLineJumpStyle",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{0,L"Shape Layout",L"",L"ConLineJumpDirX",visSectionObject,visRowShapeLayout,visSLOJumpDirX,L"ConLineJumpDirX",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{0,L"Shape Layout",L"",L"ConLineJumpDirY",visSectionObject,visRowShapeLayout,visSLOJumpDirY,L"ConLineJumpDirY",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{0,L"Shape Layout",L"",L"ShapePlaceFlip",visSectionObject,visRowShapeLayout,visSLOPlaceFlip,L"ShapePlaceFlip",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{0,L"Shape Layout",L"",L"ConLineRouteExt",visSectionObject,visRowShapeLayout,visSLOLineRouteExt,L"ConLineRouteExt",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{0,L"Shape Layout",L"",L"ShapeSplit",visSectionObject,visRowShapeLayout,visSLOSplit,L"ShapeSplit",L"0=visSLOSplitNone;1=visSLOSplitAllow"},
	{0,L"Shape Layout",L"",L"ShapeSplittable",visSectionObject,visRowShapeLayout,visSLOSplittable,L"ShapeSplittable",L"0=visSLOSplittableNone;1=visSLOSplittableAllow"},
	{14,L"Shape Layout",L"",L"DisplayLevel",visSectionObject,visRowShapeLayout,visSLODisplayLevel,L"DisplayLevel",L""},
	{0,L"Shape Transform",L"",L"PinX",visSectionObject,visRowXFormOut,visXFormPinX,L"PinX",L""},
	{0,L"Shape Transform",L"",L"PinY",visSectionObject,visRowXFormOut,visXFormPinY,L"PinY",L""},
	{0,L"Shape Transform",L"",L"Width",visSectionObject,visRowXFormOut,visXFormWidth,L"Width",L""},
	{0,L"Shape Transform",L"",L"Height",visSectionObject,visRowXFormOut,visXFormHeight,L"Height",L""},
	{0,L"Shape Transform",L"",L"LocPinX",visSectionObject,visRowXFormOut,visXFormLocPinX,L"LocPinX",L""},
	{0,L"Shape Transform",L"",L"LocPinY",visSectionObject,visRowXFormOut,visXFormLocPinY,L"LocPinY",L""},
	{0,L"Shape Transform",L"",L"Angle",visSectionObject,visRowXFormOut,visXFormAngle,L"Angle",L""},
	{0,L"Shape Transform",L"",L"FlipX",visSectionObject,visRowXFormOut,visXFormFlipX,L"FlipX",L"TRUE;FALSE"},
	{0,L"Shape Transform",L"",L"FlipY",visSectionObject,visRowXFormOut,visXFormFlipY,L"FlipY",L"TRUE;FALSE"},
	{0,L"Shape Transform",L"",L"ResizeMode",visSectionObject,visRowXFormOut,visXFormResizeMode,L"ResizeMode",L"0=visXFormResizeDontCare;1=visXFormResizeSpread;2=visXFormResizeScale"},
	{0,L"Style Properties",L"",L"EnableLineProps",visSectionObject,visRowStyle,visStyleIncludesLine,L"EnableLineProps",L"TRUE;FALSE"},
	{0,L"Style Properties",L"",L"EnableFillProps",visSectionObject,visRowStyle,visStyleIncludesFill,L"EnableFillProps",L"TRUE;FALSE"},
	{0,L"Style Properties",L"",L"EnableTextProps",visSectionObject,visRowStyle,visStyleIncludesText,L"EnableTextProps",L"TRUE;FALSE"},
	{0,L"Style Properties",L"",L"HideForApply",visSectionObject,visRowStyle,visStyleHidden,L"HideForApply",L"TRUE;FALSE"},
	{0,L"Tabs",L"Tabs.{i}{j}",L"Tabs.{i}{j}",visSectionTab,visRowTab ,visTabPos,L"Tabs.{i}{j}",L""},
	{0,L"Tabs",L"Tabs.{i}{j}",L"Tabs.{i}{j}",visSectionTab,visRowTab ,visTabAlign,L"Tabs.{i}{j}",L"0=visTabStopLeft;1=visTabStopCenter;2=visTabStopRight;3=visTabStopDecimal;4=visTabStopComma"},
	{0,L"Text Block Format",L"",L"LeftMargin",visSectionObject,visRowText,visTxtBlkLeftMargin,L"LeftMargin",L""},
	{0,L"Text Block Format",L"",L"RightMargin",visSectionObject,visRowText,visTxtBlkRightMargin,L"RightMargin",L""},
	{0,L"Text Block Format",L"",L"TopMargin",visSectionObject,visRowText,visTxtBlkTopMargin,L"TopMargin",L""},
	{0,L"Text Block Format",L"",L"BottomMargin",visSectionObject,visRowText,visTxtBlkBottomMargin,L"BottomMargin",L""},
	{0,L"Text Block Format",L"",L"VerticalAlign",visSectionObject,visRowText,visTxtBlkVerticalAlign,L"VerticalAlign",L"0=visVertTop;1=visVertMiddle;2=visVertBottom"},
	{0,L"Text Block Format",L"",L"TextBkgnd",visSectionObject,visRowText,visTxtBlkBkgnd,L"TextBkgnd",L""},
	{0,L"Text Block Format",L"",L"DefaultTabstop",visSectionObject,visRowText,visTxtBlkDefaultTabStop,L"DefaultTabstop",L""},
	{0,L"Text Block Format",L"",L"TextDirection",visSectionObject,visRowText,visTxtBlkDirection,L"TextDirection",L"0=visTxtBlkLeftToRight;1=visTxtBlkTopToBottom"},
	{0,L"Text Block Format",L"",L"TextBkgndTrans",visSectionObject,visRowText,visTxtBlkBkgndTrans,L"TextBkgndTrans",L"0-100"},
	{0,L"Text Fields",L"{i}",L"Value",visSectionTextField,visRowField ,visFieldCell,L"Fields.Value[{i}]",L""},
	{0,L"Text Fields",L"{i}",L"Format",visSectionTextField,visRowField ,visFieldFormat,L"Fields.Format[{i}]",L""},
	{0,L"Text Fields",L"{i}",L"Type",visSectionTextField,visRowField ,visFieldType,L"Fields.Type[{i}]",L"0;2;5;6;7"},
	{0,L"Text Fields",L"{i}",L"UICat",visSectionTextField,visRowField ,visFieldUICategory,L"Fields.UICat[{i}]",L""},
	{0,L"Text Fields",L"{i}",L"UICod",visSectionTextField,visRowField ,visFieldUICode,L"Fields.UICod[{i}]",L""},
	{0,L"Text Fields",L"{i}",L"UIFmt",visSectionTextField,visRowField ,visFieldUIFormat,L"Fields.UIFmt[{i}]",L""},
	{0,L"Text Fields",L"{i}",L"Calendar",visSectionTextField,visRowField ,visFieldCalendar,L"Fields.Calendar[{i}]",L"0=Western;1=Arabic Hijri;2=Hebrew Lunar;3=Taiwan Calendar;4=Japanese Emperor Reign;5=Thai Buddhist;6=Korean Danki;7=Saka Era;8=English transliterated;9=French transliterated"},
	{0,L"Text Fields",L"{i}",L"ObjectKind",visSectionTextField,visRowField ,visFieldObjectKind,L"Fields.ObjectKind[{i}]",L"0=visTFOKStandard;1=visTFOKHorizontaInVertical"},
	{0,L"Text Transform",L"",L"TxtPinX",visSectionObject,visRowTextXForm,visXFormPinX,L"TxtPinX",L""},
	{0,L"Text Transform",L"",L"TxtPinY",visSectionObject,visRowTextXForm,visXFormPinY,L"TxtPinY",L""},
	{0,L"Text Transform",L"",L"TxtWidth",visSectionObject,visRowTextXForm,visXFormWidth,L"TxtWidth",L""},
	{0,L"Text Transform",L"",L"TxtHeight",visSectionObject,visRowTextXForm,visXFormHeight,L"TxtHeight",L""},
	{0,L"Text Transform",L"",L"TxtLocPinX",visSectionObject,visRowTextXForm,visXFormLocPinX,L"TxtLocPinX",L""},
	{0,L"Text Transform",L"",L"TxtLocPinY",visSectionObject,visRowTextXForm,visXFormLocPinY,L"TxtLocPinY",L""},
	{0,L"Text Transform",L"",L"TxtAngle",visSectionObject,visRowTextXForm,visXFormAngle,L"TxtAngle",L""},
	{15,L"Theme Properties",L"",L"ColorSchemeIndex",visSectionObject,visRowThemeProperties,visColorSchemeIndex,L"ColorSchemeIndex",L""},
	{15,L"Theme Properties",L"",L"EffectSchemeIndex",visSectionObject,visRowThemeProperties,visEffectSchemeIndex,L"EffectSchemeIndex",L""},
	{15,L"Theme Properties",L"",L"ConnectorSchemeIndex",visSectionObject,visRowThemeProperties,visConnectorSchemeIndex,L"ConnectorSchemeIndex",L""},
	{15,L"Theme Properties",L"",L"FontSchemeIndex",visSectionObject,visRowThemeProperties,visFontSchemeIndex,L"FontSchemeIndex",L""},
	{15,L"Theme Properties",L"",L"ThemeIndex",visSectionObject,visRowThemeProperties,visThemeIndex,L"ThemeIndex",L""},
	{15,L"Theme Properties",L"",L"VariationColorIndex",visSectionObject,visRowThemeProperties,visVariationColorIndex,L"VariationColorIndex",L""},
	{15,L"Theme Properties",L"",L"VariationStyleIndex",visSectionObject,visRowThemeProperties,visVariationStyleIndex,L"VariationStyleIndex",L""},
	{15,L"Theme Properties",L"",L"EmbellishmentIndex",visSectionObject,visRowThemeProperties,visEmbellishmentIndex,L"EmbellishmentIndex",L""},
	{0,L"User-Defined Cells",L"{r}",L"Value",visSectionUser,visRowUser ,visUserValue,L"User.{r}.Value",L""},
	{0,L"User-Defined Cells",L"{r}",L"Prompt",visSectionUser,visRowUser ,visUserPrompt,L"User.{r}.Prompt",L""},

};

typedef std::vector<SSInfo> SSInfos;

const SSInfos& GetSectionInfo(short section)
{
	static SSInfos index[255];

	struct Index
	{
		Index()
		{
			for (size_t i = 0; i < _countof(ss_global); ++i)
			{
				SSInfo& ss_global_item = ss_global[i];
				
				ss_global_item.index = i;

				if (ss_global_item.visio_version <= GetVisioVersion())
					index[ss_global_item.s].push_back(ss_global_item);
			}
		}
	};

	static Index index_init;

	return index[section];
}

void AddNameMatchResult(const CString& mask, CString name, 
						short s, short r, short c, 
						const CString& s_name, const CString& r_name_l, const CString r_name_u, const CString& c_name, size_t index,
						std::set<SRC> &result)
{
	Strings masks;
	SplitList(mask, L",", masks);

	CString lcase_name = name;
	lcase_name.MakeLower();

	for (Strings::const_iterator it = masks.begin(); it != masks.end(); ++it)
	{
		CString mask = *it;
		mask.Trim();
		mask.MakeLower();

		if (StringIsLike(mask, lcase_name))
		{
			SRC src;

			src.name = name;
			src.s = s;
			src.s_name = s_name;
			src.r = r;
			src.r_name_l = r_name_l;
			src.r_name_u = r_name_l;
			src.c = c;
			src.c_name = c_name;
			src.index = index;

			result.insert(src);
			return;
		}
	}
}

void GetVariableIndexedSectionCellNames(IVShape* shape, short section_no, const CString& mask, std::set<SRC>& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	const SSInfos& ss_info = GetSectionInfo(section_no);

	for (short r = 0; r < section->Count; ++r)
	{
		for (size_t i = 0; i < ss_info.size(); ++i)
		{
			CString name = ss_info[i].name;

			CString r_name = FormatString(L"%d", 1 + r);
			name.Replace(L"{i}", r_name);

			AddNameMatchResult(mask, name, 
				section_no, r, ss_info[i].c, 
				ss_info[i].s_name, r_name, r_name, ss_info[i].c_name, 
				ss_info[i].index,
				result);
		}
	}
}

void GetVariableNamedSectionCellNames(IVShape* shape, short section_no, const CString& mask, std::set<SRC>& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	const SSInfos& ss_info = GetSectionInfo(section_no);

	for (short r = 0; r < section->Count; ++r)
 	{
		IVRowPtr row = section->GetRow(r);
		CString row_name = row->Name;
		CString row_name_u = row->NameU;

		for (size_t i = 0; i < ss_info.size(); ++i)
		{
			CString name = ss_info[i].name;
			name.Replace(L"{r}", row_name);

			AddNameMatchResult(mask, name, 
				section_no, r, ss_info[i].c, 
				ss_info[i].s_name, row_name, row_name_u, ss_info[i].c_name, 
				ss_info[i].index,
				result);
		}
	}
}

short GetGeometryRowCellCount(short row_type)
{
	switch (row_type)
	{
	case visTagComponent:			return 6;
	case visTagMoveTo:				return 2;
	case visTagLineTo:				return 2;
	case visTagArcTo:				return 3;
	case visTagInfiniteLine:		return 4;
	case visTagEllipse:				return 6;
	case visTagEllipticalArcTo: 	return 6;
	case visTagSplineBeg:			return 6;
	case visTagSplineSpan:			return 3;
	case visTagPolylineTo:			return 3;
	case visTagNURBSTo:				return 7;
	case visTagRelMoveTo:			return 2;
	case visTagRelLineTo:			return 2;
	case visTagRelEllipticalArcTo:	return 6;
	case visTagRelCubBezTo:			return 6;
	case visTagRelQuadBezTo:		return 4;

	default:
		return (6);
	}
}

void GetVariableGeometrySectionCellNames(IVShape* shape, const CString& mask, std::set<SRC>& result)
{
	const SSInfos& ss_infos = GetSectionInfo(visSectionFirstComponent);

	for (short s = 0; s < shape->GeometryCount; ++s)
	{
		IVSectionPtr section = shape->GetSection(visSectionFirstComponent + s);

		CString s_name = FormatString(L"%d", 1 + s);

		for (short r = 0; r < section->Count; ++r)
		{
			short c_count = GetGeometryRowCellCount(shape->GetRowType(visSectionFirstComponent, r));

			CString r_name = FormatString(L"%d", r);

			for (size_t i = 0; i < ss_infos.size(); ++i)
			{
				if (ss_infos[i].c >= c_count)
					continue;

				if (ss_infos[i].r == visRowComponent && r == 0)
				{
					CString name = ss_infos[i].name;
					name.Replace(L"{i}", s_name);

					AddNameMatchResult(mask, name, 
						visSectionFirstComponent + s, r, ss_infos[i].c, 
						s_name, L"", L"", ss_infos[i].c_name, 
						ss_infos[i].index,
						result);

					continue;
				}

				if (ss_infos[i].r == visRowVertex && r >= 1)
				{
					CString name = ss_infos[i].name;
					name.Replace(L"{i}", s_name);
					name.Replace(L"{j}", r_name);

					AddNameMatchResult(mask, name, 
						visSectionFirstComponent + s, r, ss_infos[i].c, 
						s_name, r_name, r_name, ss_infos[i].c_name, 
						ss_infos[i].index,
						result);
				}
			}
		}
	}
}

void GetSimpleSectionCellNames(IVShape* shape, const CString& mask, std::set<SRC>& result)
{
	const SSInfos& ss_info = GetSectionInfo(visSectionObject);
	for (size_t i = 0; i < ss_info.size(); ++i)
	{
		if (!shape->GetRowExists(visSectionObject, ss_info[i].r, VARIANT_FALSE))
			continue;

		CString name = ss_info[i].name;

		AddNameMatchResult(mask, name, 
			ss_info[i].s, ss_info[i].r, ss_info[i].c, 
			ss_info[i].s_name, L"", L"", ss_info[i].c_name,
			ss_info[i].index,
			result);
	}
}

bool IsNamedRowSection(short s)
{
	switch (s)
	{
	case visSectionAction:
	case visSectionSmartTag:
	case visSectionControls:
	case visSectionHyperlink:
	case visSectionProp:
	case visSectionUser:
		return true;

	default:
		return false;
	}
}

LPCWSTR GetNamedSectionName(short s)
{
	switch (s)
	{
	case visSectionAction:		return L"Actions";
	case visSectionSmartTag:	return L"SmartTags";
	case visSectionControls:	return L"Controls";
	case visSectionHyperlink:	return L"Hyperlink";
	case visSectionProp:		return L"Prop";
	case visSectionUser:		return L"User";

	default:
		ASSERT_RETURN_VALUE(FALSE, L"");
	}
}


void GetCellNames(IVShape* shape, const CString& cell_name_mask, std::set<SRC>& result)
{
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

const SSInfo& GetSSInfo(size_t index)
{
	if (0 <= index && index <= _countof(ss_global))
		return ss_global[index];

	ASSERT(FALSE);

	static SSInfo stub;
	return stub;
}

bool CellExists(IVShape* shape, const SRC& src)
{
	if (IsNamedRowSection(src.s))
	{
		return shape->GetCellExistsU(
			bstr_t(FormatString(L"%s.%s.%s", GetNamedSectionName(src.s), src.r_name_u, src.c_name)), 
			VARIANT_FALSE) != VARIANT_FALSE;
	}
	else
	{
		return shape->GetCellsSRCExists(
			src.s, src.r, src.c, 
			VARIANT_FALSE) != VARIANT_FALSE;	
	}
}

IVCellPtr GetShapeCell(IVShape* shape, const SRC& src )
{
	if (IsNamedRowSection(src.s))
	{
		return shape->GetCellsU(
			bstr_t(FormatString(L"%s.%s.%s", GetNamedSectionName(src.s), src.r_name_u, src.c_name)));
	}
	else
	{
		return shape->GetCellsSRC(src.s, src.r, src.c);	
	}
}

bool SRC::operator < (const SRC& other) const
{
	if (s < other.s) return true;
	if (s > other.s) return false;

	if (IsNamedRowSection(s))
	{
		if (r_name_u < other.r_name_u) return true;
		if (r_name_u > other.r_name_u) return false;
	}
	else
	{
		if (r < other.r) return true;
		if (r > other.r) return false;
	}

	if (c < other.c) return true;
	if (c > other.c) return false;

	return (index < other.index);
}

} // namespace shapesheet
