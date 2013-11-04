
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
	short s;
	short r;
	short c;
	LPCWSTR name;
	LPCWSTR type;
	LPCWSTR values;
};

#define countof(a)	(sizeof(a)/sizeof(a[0]))

SSInfo ss_info[] = {

	{L"1-D Endpoints",visSectionObject,visRowXForm1D,vis1DBeginX,L"BeginX",L"",L""},
	{L"1-D Endpoints",visSectionObject,visRowXForm1D,vis1DBeginY,L"BeginY",L"",L""},
	{L"1-D Endpoints",visSectionObject,visRowXForm1D,vis1DEndX,L"EndX",L"",L""},
	{L"1-D Endpoints",visSectionObject,visRowXForm1D,vis1DEndY,L"EndY",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visKeepTextFlat,L"KeepTextFlat",L"BOOL",L"TRUE;FALSE"},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visPerspective,L"Perspective",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visRotationType,L"RotationType",L"",L"0;1;2;3;4;5;6"},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visRotationXAngle,L"RotationXAngle",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visRotationYAngle,L"RotationYAngle",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visRotationZAngle,L"RotationZAngle",L"",L""},
	{L"3-D Rotation Properties",visSectionObject,visRow3DRotationProperties,visDistanceFromGround,L"DistanceFromGround",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagButtonFace,L"SmartTags.{r}.ButtonFace",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagDescription,L"SmartTags.{r}.Description",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagDisabled,L"SmartTags.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagDisplayMode,L"SmartTags.{r}.DisplayMode",L"ENUM",L"0=visSmartTagDispModeMouseOver;1=visSmartTagDispModeShapeSelected;2=visSmartTagDispModeAlways"},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagName,L"SmartTags.{r}.TagName",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagX,L"SmartTags.{r}.X",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagXJustify,L"SmartTags.{r}.XJustify",L"ENUM",L"0=visSmartTagXJustifyLeft;1=visSmartTagXJustifyCenter;2=visSmartTagXJustifyRight"},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagY,L"SmartTags.{r}.Y",L"",L""},
	{L"Action Tags",visSectionSmartTag,visRowSmartTag ,visSmartTagYJustify,L"SmartTags.{r}.YJustify",L"",L"0=visSmartTagYJustifyTop;1=visSmartTagYJustifyMiddle;2=visSmartTagYJustifyBottom"},
	{L"Actions",visSectionAction,visRowAction ,visActionAction,L"Actions.{r}.Action",L"",L""},
	{L"Actions",visSectionAction,visRowAction ,visActionBeginGroup,L"Actions.{r}.BeginGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,visRowAction ,visActionButtonFace,L"Actions.{r}.ButtonFace",L"",L""},
	{L"Actions",visSectionAction,visRowAction ,visActionChecked,L"Actions.{r}.Checked",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,visRowAction ,visActionDisabled,L"Actions.{r}.Disabled",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,visRowAction ,visActionFlyoutChild,L"Actions.{r}.FlyoutChild",L"",L""},
	{L"Actions",visSectionAction,visRowAction ,visActionInvisible,L"Actions.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,visRowAction ,visActionMenu,L"Actions.{r}.Menu",L"",L""},
	{L"Actions",visSectionAction,visRowAction ,visActionReadOnly,L"Actions.{r}.ReadOnly",L"BOOL",L"TRUE;FALSE"},
	{L"Actions",visSectionAction,visRowAction ,visActionSortKey,L"Actions.{r}.SortKey",L"",L""},
	{L"Actions",visSectionAction,visRowAction ,visActionTagName,L"Actions.{r}.TagName",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visGlowColor,L"GlowColor",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visGlowColorTrans,L"GlowColorTrans",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visGlowSize,L"GlowSize",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visReflectionBlur,L"ReflectionBlur",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visReflectionDist,L"ReflectionDist",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visReflectionSize,L"ReflectionSize",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visReflectionTrans,L"ReflectionTrans",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchAmount,L"SketchAmount",L"",L"0;1-25"},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchEnabled,L"SketchEnabled",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchFillChange,L"SketchFillChange",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchLineChange,L"SketchLineChange",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchLineWeight,L"SketchLineWeight",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSketchSeed,L"SketchSeed",L"",L""},
	{L"Additional Effect Properties",visSectionObject,visRowOtherEffectProperties,visSoftEdgesSize,L"SoftEdgesSize",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignBottom,L"AlignBottom",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignCenter,L"AlignCenter",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignLeft,L"AlignLeft",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignMiddle,L"AlignMiddle",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignRight,L"AlignRight",L"",L""},
	{L"Alignment",visSectionObject,visRowAlign,visAlignTop,L"AlignTop",L"",L""},
	{L"Annotation",visSectionAnnotation,visRowAnnotation ,visAnnotationComment,L"Annotation.Comment[{i}]",L"",L""},
	{L"Annotation",visSectionAnnotation,visRowAnnotation ,visAnnotationDate,L"Annotation.Date[{i}]",L"",L""},
	{L"Annotation",visSectionAnnotation,visRowAnnotation ,visAnnotationLangID,L"Annotation.LangID[{i}]",L"",L""},
	{L"Annotation",visSectionAnnotation,visRowAnnotation ,visAnnotationX,L"Annotation.X[{i}]",L"",L""},
	{L"Annotation",visSectionAnnotation,visRowAnnotation ,visAnnotationY,L"Annotation.Y[{i}]",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelBottomHeight,L"BevelBottomHeight",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelBottomType,L"BevelBottomType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelBottomWidth,L"BevelBottomWidth",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelContourColor,L"BevelContourColor",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelContourSize,L"BevelContourSize",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelDepthColor,L"BevelDepthColor",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelDepthSize,L"BevelDepthSize",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelLightingAngle,L"BevelLightingAngle",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelLightingType,L"BevelLightingType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12;13;14;15"},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelMaterialType,L"BevelMaterialType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11"},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelTopHeight,L"BevelTopHeight",L"",L""},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelTopType,L"BevelTopType",L"",L"0;1;2;3;4;5;6;7;8;9;10;11;12"},
	{L"Bevel Properties",visSectionObject,visRowBevelProperties,visBevelTopWidth,L"BevelTopWidth",L"",L""},
	{L"Change Shape Behavior",visSectionObject,visRowReplaceBehaviors,visReplaceCopyCells,L"ReplaceCopyCells",L"",L""},
	{L"Change Shape Behavior",visSectionObject,visRowReplaceBehaviors,visReplaceLockFormat,L"ReplaceLockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",visSectionObject,visRowReplaceBehaviors,visReplaceLockShapeData,L"ReplaceLockShapeData",L"BOOL",L"TRUE;FALSE"},
	{L"Change Shape Behavior",visSectionObject,visRowReplaceBehaviors,visReplaceLockText,L"ReplaceLockText",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterAsianFont,L"Char.AsianFont[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterCase,L"Char.Case[{i}]",L"ENUM",L"0=visCaseNormal;1=visCaseAllCaps;2=visCaseInitialCaps"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterColor,L"Char.Color[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterColorTrans,L"Char.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptFont,L"Char.ComplexScriptFont[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterComplexScriptSize,L"Char.ComplexScriptSize[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterDblUnderline,L"Char.DblUnderline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterDoubleStrikethrough,L"Char.DoubleStrikethrough[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterFont,L"Char.Font[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterFontScale,L"Char.FontScale[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterLangID,L"Char.LangID[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterLetterspace,L"Char.Letterspace[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterOverline,L"Char.Overline[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterPos,L"Char.Pos[{i}]",L"ENUM",L"0=visPosNormal;1=visPosSuper;2=visPosSub"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterSize,L"Char.Size[{i}]",L"",L""},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterStrikethru,L"Char.Strikethru[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Character",visSectionCharacter,visRowCharacter ,visCharacterStyle,L"Char.Style[{i}]",L"",L"Bold=visBold;Italic=visItalic;Underline=visUnderLine;Small caps=visSmallCaps"},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visCnnctD,L"Connections.D[{i}]",L"",L""},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visCnnctA,L"Connections.DirX[{i}]",L"",L""},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visCnnctC,L"Connections.Type[{i}]",L"ENUM",L"0=visCnnctTypeInward;1=visCnnctTypeOutward;2=visCnnctTypeInwardOutward"},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visX,L"Connections.X{i}",L"",L""},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visY,L"Connections.Y{i}",L"",L""},
	{L"Connection Points",visSectionConnectionPts,visRowConnectionPts ,visCnnctB,L"Connections.DirY[{i}]",L"",L""},
	{L"Controls",visSectionControls,visRowControl ,visCtlGlue,L"Controls.{r}.CanGlue",L"BOOL",L"TRUE;FALSE"},
	{L"Controls",visSectionControls,visRowControl ,visCtlTip,L"Controls.{r}.Tip",L"",L""},
	{L"Controls",visSectionControls,visRowControl ,visCtlX,L"Controls.{r}.X",L"",L""},
	{L"Controls",visSectionControls,visRowControl ,visCtlXCon,L"Controls.{r}.XCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",visSectionControls,visRowControl ,visCtlXDyn,L"Controls.{r}.XDyn",L"",L""},
	{L"Controls",visSectionControls,visRowControl ,visCtlY,L"Controls.{r}.Y",L"",L""},
	{L"Controls",visSectionControls,visRowControl ,visCtlYCon,L"Controls.{r}.YCon",L"ENUM",L"0=visCtlProportional;1=visCtlLocked;2=visCtlOffsetMin;3=visCtlOffsetMid;4=visCtlOffsetMax;5=visCtlProportionalHidden;6=visCtlLockedHiddenv;7=visCtlOffsetMinHidden;8=visCtlOffsetMidHidden;9=visCtlOffsetMaxHidden"},
	{L"Controls",visSectionControls,visRowControl ,visCtlYDyn,L"Controls.{r}.YDyn",L"",L""},
	{L"Document Properties",visSectionObject,visRowDoc,visDocAddMarkup,L"AddMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocLangID,L"DocLangID",L"",L""},
	{L"Document Properties",visSectionObject,visRowDoc,visDocLockDuplicatePage,L"DocLockDuplicatePage",L"",L""},
	{L"Document Properties",visSectionObject,visRowDoc,visDocLockReplace,L"DocLocReplace",L"",L""},
	{L"Document Properties",visSectionObject,visRowDoc,visDocLockPreview,L"LockPreview",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocNoCoauth,L"NoCoauth",L"BOOL",L"TRUE;FALSE"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocOutputFormat,L"OutputFormat",L"ENUM",L"0;1;2"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocPreviewQuality,L"PreviewQuality",L"ENUM",L"0=visDocPreviewQualityDraft;1=visDocPreviewQualityDetailed"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocPreviewScope,L"PreviewScope",L"ENUM",L"0=visDocPreviewScope1stPage;1=visDocPreviewScopeNone;2=visDocPreviewScopeAllPages"},
	{L"Document Properties",visSectionObject,visRowDoc,visDocViewMarkup,L"ViewMarkup",L"BOOL",L"TRUE;FALSE"},
	{L"Events",visSectionObject,visRowEvent,visEvtCellDblClick,L"EventDblClick",L"",L""},
	{L"Events",visSectionObject,visRowEvent,visEvtCellDrop,L"EventDrop",L"",L""},
	{L"Events",visSectionObject,visRowEvent,visEvtCellMultiDrop,L"EventMultiDrop",L"",L""},
	{L"Events",visSectionObject,visRowEvent,visEvtCellXFMod,L"EventXFMod",L"",L""},
	{L"Events",visSectionObject,visRowEvent,visEvtCellTheText,L"TheText",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillBkgnd,L"FillBkgnd",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillBkgndTrans,L"FillBkgndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",visSectionObject,visRowFill,visFillForegnd,L"FillForegnd",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillForegndTrans,L"FillForegndTrans",L"ENUM",L"0 - 100"},
	{L"Fill Format",visSectionObject,visRowFill,visFillPattern,L"FillPattern",L"ENUM",L"0;1;2 - 40"},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwBlur,L"ShapeShdwBlur",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwObliqueAngle,L"ShapeShdwObliqueAngle",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwOffsetX,L"ShapeShdwOffsetX",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwOffsetY,L"ShapeShdwOffsetY",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwScaleFactor,L"ShapeShdwScaleFactor",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwShow,L"ShapeShdwShow",L"ENUM",L"0;1;2"},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwType,L"ShapeShdwType",L"ENUM",L"0=visFSTPageDefault;1=visFSTSimple;2=visFSTOblique"},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwForegnd,L"ShdwForegnd",L"",L""},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwForegndTrans,L"ShdwForegndTrans",L"",L"0 - 100"},
	{L"Fill Format",visSectionObject,visRowFill,visFillShdwPattern,L"ShdwPattern",L"",L"0;1;2 - 40"},
	{L"Foreign Image Info",visSectionObject,visRowForeign,visFrgnImgClippingPath,L"ClippingPath",L"",L""},
	{L"Foreign Image Info",visSectionObject,visRowForeign,visFrgnImgHeight,L"ImgHeight",L"",L""},
	{L"Foreign Image Info",visSectionObject,visRowForeign,visFrgnImgOffsetX,L"ImgOffsetX",L"",L""},
	{L"Foreign Image Info",visSectionObject,visRowForeign,visFrgnImgOffsetY,L"ImgOffsetY",L"",L""},
	{L"Foreign Image Info",visSectionObject,visRowForeign,visFrgnImgWidth,L"ImgWidth",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,2,L"Geometry{i}.A{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,3,L"Geometry{i}.B{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,4,L"Geometry{i}.C{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,5,L"Geometry{i}.D{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,visNURBSData,L"Geometry{i}.E{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowComponent,visCompNoFill,L"Geometry{i}.NoFill",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,visRowComponent,visCompNoLine,L"Geometry{i}.NoLine",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,visRowComponent,visCompNoQuickDrag,L"Geometry{i}.NoQuickDrag",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowComponent,visCompNoShow,L"Geometry{i}.NoShow",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,visRowComponent,visCompNoSnap,L"Geometry{i}.NoSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,visX,L"Geometry{i}.X{j}",L"",L""},
	{L"Geometry",visSectionFirstComponent,visRowVertex ,visY,L"Geometry{i}.Y{j}",L"",L""},
	{L"Glue Info",visSectionObject,visRowGroup,visBegTrigger,L"BegTrigger",L"",L""},
	{L"Glue Info",visSectionObject,visRowMisc,visEndTrigger,L"EndTrigger",L"",L""},
	{L"Glue Info",visSectionObject,visRowMisc,visGlueType,L"GlueType",L"ENUM",L"0=visGlueTypeDefault;2=visGlueTypeWalking;4=visGlueTypeNoWalking;8=visGlueTypeNoWalkingTo"},
	{L"Glue Info",visSectionObject,visRowMisc,visWalkPref,L"WalkPreference",L"ENUM",L"1=visWalkPrefBegNS;2=visWalkPrefEndNS"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visFillGradientAngle,L"FillGradientAngle",L"",L""},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visFillGradientDir,L"FillGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visFillGradientEnabled,L"FillGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visLineGradientAngle,L"LineGradientAngle",L"",L""},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visLineGradientDir,L"LineGradientDir",L"",L"0;1-7;8-12;13"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visLineGradientEnabled,L"LineGradientEnabled",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visRotateGradientWithShape,L"RotateGradientWithShape",L"BOOL",L"TRUE;FALSE"},
	{L"Gradient Properties",visSectionObject,visRowGradientProperties,visUseGroupGradient,L"UseGroupGradient",L"",L""},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupDisplayMode,L"DisplayMode",L"ENUM",L"0=visGrpDispModeNone;1=visGrpDispModeBack;2=visGrpDispModeFront"},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupDontMoveChildren,L"DontMoveChildren",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupIsDropTarget,L"IsDropTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupIsSnapTarget,L"IsSnapTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupIsTextEditTarget,L"IsTextEditTarget",L"BOOL",L"TRUE;FALSE"},
	{L"Group Properties",visSectionObject,visRowGroup,visGroupSelectMode,L"SelectMode",L"",L"0=visGrpSelModeGroupOnly;1=visGrpSelModeGroup1st;2=visGrpSelModeMembers1st"},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkAddress,L"Hyperlink.{r}.Address",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkDefault,L"Hyperlink.{r}.Default",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkDescription,L"Hyperlink.{r}.Description",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkExtraInfo,L"Hyperlink.{r}.ExtraInfo",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkFrame,L"Hyperlink.{r}.Frame",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkInvisible,L"Hyperlink.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkNewWin,L"Hyperlink.{r}.NewWindow",L"BOOL",L"TRUE;FALSE"},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkSortKey,L"Hyperlink.{r}.SortKey",L"",L""},
	{L"Hyperlinks",visSectionHyperlink,visRow1stHyperlink ,visHLinkSubAddress,L"Hyperlink.{r}.SubAddress",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageBlur,L"Blur",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageBrightness,L"Brightness",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageContrast,L"Contrast",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageDenoise,L"Denoise",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageGamma,L"Gamma",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageSharpen,L"Sharpen",L"",L""},
	{L"Image Properties",visSectionObject,visRowImage,visImageTransparency,L"Transparency",L"",L"0 - 100"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerActive,L"Layers.Active[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerColor,L"Layers.Color[{i}]",L"",L""},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerColorTrans,L"Layers.ColorTrans[{i}]",L"",L"0 - 100"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerGlue,L"Layers.Glue[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerLock,L"Layers.Locked[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,visRowLayer ,visDocPreviewScope,L"Layers.Print[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerSnap,L"Layers.Snap[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Layers",visSectionLayer,visRowLayer ,visLayerVisible,L"Layers.Visible[{i}]",L"BOOL",L"TRUE;FALSE"},
	{L"Line Format",visSectionObject,visRowLine,visLineBeginArrow,L"BeginArrow",L"",L"0;1 - 45"},
	{L"Line Format",visSectionObject,visRowLine,visLineBeginArrowSize,L"BeginArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",visSectionObject,visRowLine,visCompoundType,L"CompoundType",L"",L"0;1;2;3;4"},
	{L"Line Format",visSectionObject,visRowLine,visLineEndArrow,L"EndArrow",L"",L"0;1 - 45"},
	{L"Line Format",visSectionObject,visRowLine,visLineEndArrowSize,L"EndArrowSize",L"",L"0=visArrowSizeVerySmall;1=visArrowSizeSmall;2=visArrowSizeMedium;3=visArrowSizeLarge;4=visArrowSizeVeryLarge;5=visArrowSizeJumbo;6=visArrowSizeColossal"},
	{L"Line Format",visSectionObject,visRowLine,visLineEndCap,L"LineCap",L"",L"0;1;2"},
	{L"Line Format",visSectionObject,visRowLine,visLineColor,L"LineColor",L"",L""},
	{L"Line Format",visSectionObject,visRowLine,visLineColorTrans,L"LineColorTrans",L"",L"0 - 100"},
	{L"Line Format",visSectionObject,visRowLine,visLinePattern,L"LinePattern",L"",L"0;1;2 - 23"},
	{L"Line Format",visSectionObject,visRowLine,visLineWeight,L"LineWeight",L"",L""},
	{L"Line Format",visSectionObject,visRowLine,visLineRounding,L"Rounding",L"",L""},
	{L"Miscellaneous",visSectionObject,visRowMisc,visObjCalendar,L"Calendar",L"",L""},
	{L"Miscellaneous",visSectionObject,visRowMisc,visComment,L"Comment",L"",L""},
	{L"Miscellaneous",visSectionObject,visRowMisc,visObjDropOnPageScale,L"DropOnPageScale",L"",L""},
	{L"Miscellaneous",visSectionObject,visRowMisc,visDynFeedback,L"DynFeedback",L"",L"0=visDynFBDefault;1=visDynFBUCon3Leg;2=visDynFBUCon5Leg"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visHideText,L"HideText",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visDropSource,L"IsDropSource",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visObjLangID,L"LangID",L"",L""},
	{L"Miscellaneous",visSectionObject,visRowMisc,visObjLocalizeMerge,L"LocalizeMerge",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visNoAlignBox,L"NoAlignBox",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visNoCtlHandles,L"NoCtlHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visNoLiveDynamics,L"NoLiveDynamics",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visNonPrinting,L"NonPrinting",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visNoObjHandles,L"NoObjHandles",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visObjNoProofing,L"NoProofing",L"BOOL",L"TRUE;FALSE"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visLOFlags,L"ObjType",L"",L"0=visLOFlagsVisDecides;1=visLOFlagsPlacable;2=visLOFlagsRoutable;4=visLOFlagsDont;8=visLOFlagsPNRGroup"},
	{L"Miscellaneous",visSectionObject,visRowMisc,visUpdateAlignBox,L"UpdateAlignBox",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOAvenueSizeX,L"AvenueSizeY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOAvenueSizeY,L"AvenueSizeY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOAvoidPageBreaks,L"AvoidPageBreaks",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOBlockSizeX,L"BlockSizeX",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOBlockSizeY,L"BlockSizeY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOCtrlAsInput,L"CtrlAsInput",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLODynamicsOff,L"DynamicsOff",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOEnableGrid,L"EnableGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineAdjustFrom,L"LineAdjustFrom",L"",L"0=visPLOLineAdjustFromNotRelated;1=visPLOLineAdjustFromAll;2=visPLOLineAdjustFromNone;3=visPLOLineAdjustFromRoutingDefault"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineAdjustTo,L"LineAdjustTo",L"",L"0=visPLOLineAdjustToDefault;1=visPLOLineAdjustToAll;2=visPLOLineAdjustToNone;3=visPLOLineAdjustToRelated"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpCode,L"LineJumpCode",L"",L"0=visPLOJumpNone;1=visPLOJumpHorizontal;2=visPLOJumpVertical;3=visPLOJumpLastRouted;4=visPLOJumpDisplayOrder;5=visPLOJumpReverseDisplayOrder"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpFactorX,L"LineJumpFactorX",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpFactorY,L"LineJumpFactorY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpStyle,L"LineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineRouteExt,L"LineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineToLineX,L"LineToLineX",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineToLineY,L"LineToLineY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineToNodeX,L"LineToNodeX",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOLineToNodeY,L"LineToNodeY",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpDirX,L"PageLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOJumpDirY,L"PageLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOSplit,L"PageShapeSplit",L"",L"0=visPLOSplitNone;1=visPLOSplitAllow"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOPlaceDepth,L"PlaceDepth",L"",L"0=visPLOPlaceDepthDefault;1=visPLOPlaceDepthMedium;2=visPLOPlaceDepthDeep;3=visPLOPlaceDepthShallow"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOPlaceFlip,L"PlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOPlaceStyle,L"PlaceStyle",L"",L""},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOPlowCode,L"PlowCode",L"",L"0=visPLOPlowNone;1=visPLOPlowAll"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLOResizePage,L"ResizePage",L"BOOL",L"TRUE;FALSE"},
	{L"Page Layout",visSectionObject,visRowPageLayout,visPLORouteStyle,L"RouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Page Properties",visSectionObject,visRowPage,visPageDrawResizeType,L"DrawingResizeType",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageDrawingScale,L"DrawingScale",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageDrawScaleType,L"DrawingScaleType",L"",L"0=visNoScale;1=visArchitectural;2=visEngineering;3=visScaleCustom;4=visScaleMetric;5=visScaleMechanical"},
	{L"Page Properties",visSectionObject,visRowPage,visPageDrawSizeType,L"DrawingSizeType",L"",L"0=visPrintSetup;1=visTight;2=visStandard;3=visCustom;4=visLogical;5=visDSMetric;6=visDSEngr;7=visDSArch"},
	{L"Page Properties",visSectionObject,visRowPage,visPageInhibitSnap,L"InhibitSnap",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",visSectionObject,visRowPage,visPageHeight,L"PageHeight",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageLockDuplicate,L"PageLockDuplicate",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageLockReplace,L"PageLockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Page Properties",visSectionObject,visRowPage,visPageScale,L"PageScale",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageWidth,L"PageWidth",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageShdwObliqueAngle,L"ShdwObliqueAngle",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageShdwOffsetX,L"ShdwOffsetX",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageShdwOffsetY,L"ShdwOffsetY",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageShdwScaleFactor,L"ShdwScaleFactor",L"",L""},
	{L"Page Properties",visSectionObject,visRowPage,visPageShdwType,L"ShdwType",L"",L"1=visFSTSimple;2=visFSTOblique;3=visFSTInner"},
	{L"Page Properties",visSectionObject,visRowPage,visPageUIVisibility,L"UIVisibility",L"",L"0=visUIVNormal;1=visUIVHidden"},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visBulletIndex,L"Para.Bullet[{i}]",L"",L"0;1;2;3;4;5;6;7"},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visBulletFont,L"Para.BulletFont[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visBulletFontSize,L"Para.BulletFontSize[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visBulletString,L"Para.BulletStr[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visFlags,L"Para.Flags[{i}]",L"",L"0;1"},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visHorzAlign,L"Para.HorzAlign[{i}]",L"",L"0=visHorzLeft;1=visHorzCenter;2=visHorzRight;3=visHorzJustify;4=visHorzForce"},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visIndentFirst,L"Para.IndFirst[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visIndentLeft,L"Para.IndLeft[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visIndentRight,L"Para.IndRight[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visSpaceAfter,L"Para.SpAfter[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visSpaceBefore,L"Para.SpBefore[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visSpaceLine,L"Para.SpLine[{i}]",L"",L""},
	{L"Paragraph",visSectionParagraph,visRowParagraph ,visTextPosAfterBullet,L"Para.TextPosAfterBullet[{i}]",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterX,L"CenterX",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesCenterY,L"CenterY",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesOnPage,L"OnPage",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesBottomMargin,L"PageBottomMargin",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesLeftMargin,L"PageLeftMargin",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesRightMargin,L"PageRightMargin",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesX,L"PagesX",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesPagesY,L"PagesY",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesTopMargin,L"PageTopMargin",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperKind,L"PaperKind",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesPrintGrid,L"PrintGrid",L"BOOL",L"TRUE;FALSE"},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesPageOrientation,L"PrintPageOrientation",L"",L"0=visPPOSameAsPrinter;1=visPPOPortrait;2=visPPOLandscape"},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleX,L"ScaleX",L"",L""},
	{L"Print Properties",visSectionObject,visRowPrintProperties,visPrintPropertiesScaleY,L"ScaleY",L"",L""},
	{L"PrintProperties",visSectionObject,visRowPrintProperties,visPrintPropertiesPaperSource,L"PaperSource",L"",L""},
	{L"Protection",visSectionObject,visRowLock,visLockAspect,L"LockAspect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockBegin,L"LockBegin",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockCalcWH,L"LockCalcWH",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockCrop,L"LockCrop",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockCustProp,L"LockCustProp",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockDelete,L"LockDelete",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockEnd,L"LockEnd",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockFormat,L"LockFormat",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockFromGroupFormat,L"LockFromGroupFormat",L"",L""},
	{L"Protection",visSectionObject,visRowLock,visLockGroup,L"LockGroup",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockHeight,L"LockHeight",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockMoveX,L"LockMoveX",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockMoveY,L"LockMoveY",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockReplace,L"LockReplace",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockRotate,L"LockRotate",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockSelect,L"LockSelect",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockTextEdit,L"LockTextEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockThemeColors,L"LockThemeColors",L"",L""},
	{L"Protection",visSectionObject,visRowLock,visLockThemeConnectors,L"LockThemeConnectors",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockThemeEffects,L"LockThemeEffects",L"",L""},
	{L"Protection",visSectionObject,visRowLock,visLockThemeFonts,L"LockThemeFonts",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockThemeIndex,L"LockThemeIndex",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockVariation,L"LockVariation",L"",L""},
	{L"Protection",visSectionObject,visRowLock,visLockVtxEdit,L"LockVtxEdit",L"BOOL",L"TRUE;FALSE"},
	{L"Protection",visSectionObject,visRowLock,visLockWidth,L"LockWidth",L"BOOL",L"TRUE;FALSE"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleEffectsMatrix,L"QuickStyleEffectsMatrix",L"",L""},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillColor,L"QuickStyleFillColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleFillMatrix,L"QuickStyleFillMatrix",L"",L""},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontColor,L"QuickStyleFontColor",L"",L"0;1"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleFontMatrix,L"QuickStyleFontMatrix",L"",L""},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineColor,L"QuickStyleLineColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleLineMatrix,L"QuickStyleLineMatrix",L"",L""},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleShadowColor,L"QuickStyleShadowColor",L"",L"0;1;2;3;4;5;6;7"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleType,L"QuickStyleType",L"",L"0;1;2;3"},
	{L"Quick Style",visSectionObject,visRowQuickStyleProperties,visQuickStyleVariation,L"QuickStyleVariation",L"",L"0;1;2;4;8"},
	{L"Reviewer",visSectionReviewer,visRowReviewer ,visReviewerColor,L"Reviewer.Color[{i}]",L"",L""},
	{L"Reviewer",visSectionReviewer,visRowReviewer ,visReviewerInitials,L"Reviewer.Initials[{i}]",L"",L""},
	{L"Reviewer",visSectionReviewer,visRowReviewer ,visReviewerName,L"Reviewer.Name[{i}]",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visXGridDensity,L"XGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visXGridOrigin,L"XGridOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visXGridSpacing,L"XGridSpacing",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visXRulerDensity,L"XRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visXRulerOrigin,L"XRulerOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visYGridDensity,L"YGridDensity",L"",L"0=visGridFixed;2=visGridCoarse;4=visGridNormal;8=visGridFine"},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visYGridOrigin,L"YGridOrigin",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visYGridSpacing,L"YGridSpacing",L"",L""},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visYRulerDensity,L"YRulerDensity",L"",L"0=visRulerFixed;8=visRulerCoarse;16=visRulerNormal;32=visRulerFine"},
	{L"Ruler & Grid",visSectionObject,visRowRulerGrid,visYRulerOrigin,L"YRulerOrigin",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsAsk,L"Prop.{r}.Ask",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsCalendar,L"Prop.{r}.Calendar",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsFormat,L"Prop.{r}.Format",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsInvis,L"Prop.{r}.Invisible",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsLabel,L"Prop.{r}.Label",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsLangID,L"Prop.{r}.LangID",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsPrompt,L"Prop.{r}.Prompt",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsSortKey,L"Prop.{r}.SortKey",L"",L""},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsType,L"Prop.{r}.Type",L"",L"0=visPropTypeString;1=visPropTypeListFix;2=visPropTypeNumber;3=visPropTypeBool;4=visPropTypeListVar;5=visPropTypeDate;6=visPropTypeDuration;7=visPropTypeCurrency"},
	{L"Shape Data",visSectionProp,visRowProp ,visCustPropsValue,L"Prop.{r}.Value",L"",L""},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOConFixedCode,L"ConFixedCode",L"",L"0=visSLOConFixedRerouteFreely;1=visSLOConFixedRerouteAsNeeded;2=visSLOConFixedRerouteNever;3=visSLOConFixedRerouteOnCrossover;4=visSLOConFixedByAlgFrom;5=visSLOConFixedByAlgTo;6=visSLOConFixedByAlgFromTo"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOJumpCode,L"ConLineJumpCode",L"",L"0=visSLOJumpDefault;1=visSLOJumpNever;2=visSLOJumpAlways;3=visSLOJumpOther;4=visSLOJumpNeither"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOJumpDirX,L"ConLineJumpDirX",L"",L"0=visLOJumpDirXDefault;1=visLOJumpDirXUp;2=visLOJumpDirXDown"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOJumpDirY,L"ConLineJumpDirY",L"",L"0=visLOJumpDirYDefault;1=visLOJumpDirYLeft;2=visLOJumpDirYRight"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOJumpStyle,L"ConLineJumpStyle",L"",L"0=visLOJumpStyleDefault;1=visLOJumpStyleArc;2=visLOJumpStyleGap;3=visLOJumpStyleSquare;4=visLOJumpStyleTriangle;5=visLOJumpStyle2Point;6=visLOJumpStyle3Point;7=visLOJumpStyle4Point;8=visLOJumpStyle5Point;9=visLOJumpStyle6Point"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOLineRouteExt,L"ConLineRouteExt",L"",L"0=visLORouteExtDefault;1=visLORouteExtStraight;2=visLORouteExtNURBS"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLODisplayLevel,L"DisplayLevel",L"",L""},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLORelationships,L"Relationships",L"",L""},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOFixedCode,L"ShapeFixedCode",L"",L"1=visSLOFixedPlacement;2=visSLOFixedPlow;4=visSLOFixedPermeablePlow;32=visSLOFixedConnPtsIgnore;64=visSLOFixedConnPtsOnly;128=visSLOFixedNoFoldToShape"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPermeablePlace,L"ShapePermeablePlace",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPermX,L"ShapePermeableX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPermY,L"ShapePermeableY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPlaceFlip,L"ShapePlaceFlip",L"",L"0=visLOFlipDefault;1=visLOFlipX;2=visLOFlipY;4=visLOFlipRotate;8=visLOFlipNone"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPlaceStyle,L"ShapePlaceStyle",L"",L"visLOPlaceBottomToTop;visLOPlaceCircular;visLOPlaceCompactDownLeft;visLOPlaceCompactDownRight;visLOPlaceCompactLeftDown;visLOPlaceCompactLeftUp;visLOPlaceCompactRightDown;visLOPlaceCompactRightUp;visLOPlaceCompactUpLeft;visLOPlaceCompactUpRight;visLOPlaceDefault;visLOPlaceHierarchyBottomToTopCenter;visLOPlaceHierarchyBottomToTopLeft;visLOPlaceHierarchyBottomToTopRight;visLOPlaceHierarchyLeftToRightBottom;visLOPlaceHierarchyLeftToRightMiddle;visLOPlaceHierarchyLeftToRightTop;visLOPlaceHierarchyRightToLeftBottom;visLOPlaceHierarchyRightToLeftMiddle;visLOPlaceHierarchyRightToLeftTop;visLOPlaceHierarchyTopToBottomCenter;visLOPlaceHierarchyTopToBottomLeft;visLOPlaceHierarchyTopToBottomRight;visLOPlaceLeftToRight;visLOPlaceParentDefault;visLOPlaceRadial;visLOPlaceRightToLeft;visLOPlaceTopToBottom"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOPlowCode,L"ShapePlowCode",L"",L"0=visSLOPlowDefault;1=visSLOPlowNever;2=visSLOPlowAlways"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLORouteStyle,L"ShapeRouteStyle",L"",L"0=visLORouteDefault;1=visLORouteRightAngle;2=visLORouteStraight;3=visLORouteOrgChartNS;4=visLORouteOrgChartWE;5=visLORouteFlowchartNS;6=visLORouteFlowchartWE;7=visLORouteTreeNS;8=visLORouteTreeWE;9=visLORouteNetwork;10=visLORouteOrgChartSN;11=visLORouteOrgChartEW;12=visLORouteFlowchartSN;13=visLORouteFlowchartEW;14=visLORouteTreeSN;15=visLORouteTreeEW;16=visLORouteCenterToCenter;17=visLORouteSimpleNS;18=visLORouteSimpleWE;19=visLORouteSimpleSN;20=visLORouteSimpleEW;21=visLORouteSimpleHV;22=visLORouteSimpleVH"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOSplit,L"ShapeSplit",L"",L"0=visSLOSplitNone;1=visSLOSplitAllow"},
	{L"Shape Layout",visSectionObject,visRowShapeLayout,visSLOSplittable,L"ShapeSplittable",L"",L"0=visSLOSplittableNone;1=visSLOSplittableAllow"},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormAngle,L"Angle",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormFlipX,L"FlipX",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormFlipY,L"FlipY",L"BOOL",L"TRUE;FALSE"},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormHeight,L"Height",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormLocPinX,L"LocPinX",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormLocPinY,L"LocPinY",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormPinX,L"PinX",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormPinY,L"PinY",L"",L""},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormResizeMode,L"ResizeMode",L"",L"0=visXFormResizeDontCare;1=visXFormResizeSpread;2=visXFormResizeScale"},
	{L"Shape Transform",visSectionObject,visRowXFormOut,visXFormWidth,L"Width",L"",L""},
	{L"Style Properties",visSectionObject,visRowStyle,visStyleIncludesFill,L"EnableFillProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,visRowStyle,visStyleIncludesLine,L"EnableLineProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,visRowStyle,visStyleIncludesText,L"EnableTextProps",L"BOOL",L"TRUE;FALSE"},
	{L"Style Properties",visSectionObject,visRowStyle,visStyleHidden,L"HideForApply",L"BOOL",L"TRUE;FALSE"},
	{L"Tabs",visSectionTab,visRowTab, visTabAlign,L"Tabs.{i}{j}",L"",L"0=visTabStopLeft;1=visTabStopCenter;2=visTabStopRight;3=visTabStopDecimal;4=visTabStopComma"},
	{L"Tabs",visSectionTab,visRowTab, visTabPos,L"Tabs.{i}{j}",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkBottomMargin,L"BottomMargin",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkDefaultTabStop,L"DefaultTabstop",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkLeftMargin,L"LeftMargin",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkRightMargin,L"RightMargin",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkBkgnd,L"TextBkgnd",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkBkgndTrans,L"TextBkgndTrans",L"",L"0 - 100"},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkDirection,L"TextDirection",L"",L"0=visTxtBlkLeftToRight;1=visTxtBlkTopToBottom"},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkTopMargin,L"TopMargin",L"",L""},
	{L"Text Block Format",visSectionObject,visRowText,visTxtBlkVerticalAlign,L"VerticalAlign",L"",L"0=visVertTop;1=visVertMiddle;2=visVertBottom"},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldCalendar,L"Fields.Calendar[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldFormat,L"Fields.Format[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldObjectKind,L"Fields.ObjectKind[{i}]",L"",L"0=visTFOKStandard;1=visTFOKHorizontaInVertical"},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldType,L"Fields.Type[{i}]",L"",L"0;2;5;6;7"},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldUICategory,L"Fields.UICat[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldUICode,L"Fields.UICod[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldUIFormat,L"Fields.UIFmt[{i}]",L"",L""},
	{L"Text Fields",visSectionTextField,visRowField ,visFieldCell,L"Fields.Value[{i}]",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormAngle,L"TxtAngle",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormHeight,L"TxtHeight",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormLocPinX,L"TxtLocPinX",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormLocPinY,L"TxtLocPinY",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormPinX,L"TxtPinX",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormPinY,L"TxtPinY",L"",L""},
	{L"Text Transform",visSectionObject,visRowTextXForm,visXFormWidth,L"TxtWidth",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visColorSchemeIndex,L"ColorSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visConnectorSchemeIndex,L"ConnectorSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visEffectSchemeIndex,L"EffectSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visEmbellishmentIndex,L"EmbellishmentIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visFontSchemeIndex,L"FontSchemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visThemeIndex,L"ThemeIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visVariationColorIndex,L"VariationColorIndex",L"",L""},
	{L"Theme Properties",visSectionObject,visRowThemeProperties,visVariationStyleIndex,L"VariationStyleIndex",L"",L""},
	{L"User-Defined Cells",visSectionUser,visRowUser ,visUserPrompt,L"User.{r}.Prompt",L"",L""},
	{L"User-Defined Cells",visSectionUser,visRowUser ,visUserValue,L"User.{r}.Value",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchX,L"Scratch.X{i}",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchY,L"Scratch.Y{i}",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchA,L"Scratch.A{i}",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchB,L"Scratch.B{i}",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchC,L"Scratch.C{i}",L"",L""},
	{L"Scratch",visSectionScratch,visRowScratch,visScratchD,L"Scratch.D{i}",L"",L""},

};

void GetVariableIndexedSectionCellNames(IVShapePtr shape, short section_no, const CString& mask, std::vector<SRC>& result)
{
	if (!shape->GetSectionExists(section_no, VARIANT_FALSE))
		return;

	IVSectionPtr section = shape->GetSection(section_no);

	for (short r = 0; r < section->Count; ++r)
	{
		for (size_t i = 0; i < countof(ss_info); ++i)
		{
			if (section_no == ss_info[i].s)
			{
				CString name = ss_info[i].name;
				name.Replace(L"{i}", FormatString(L"%d", 1 + r));

				if (StringIsLike(mask, name))
				{
					SRC src;

					src.name = name;
					src.s = section_no;
					src.r = r;
					src.c = ss_info[i].c;

					result.push_back(src);
				}
			}
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

		for (size_t i = 0; i < countof(ss_info); ++i)
		{
			if (section_no == ss_info[i].s)
			{
				CString name = ss_info[i].name;
				name.Replace(L"{r}", row_name);

				if (StringIsLike(mask, name))
				{
					SRC src;

					src.name = name;
					src.s = section_no;
					src.r = r;
					src.c = ss_info[i].c;

					result.push_back(src);
				}
			}
		}
	}
}

void GetVariableGeometrySectionCellNames(IVShapePtr shape, const CString& mask, std::vector<SRC>& result)
{
	for (short s = 0; s < shape->GeometryCount; ++s)
	{
		IVSectionPtr section = shape->GetSection(visSectionFirstComponent + s);

		for (short r = 0; r < section->Count; ++r)
		{
			for (size_t i = 0; i < countof(ss_info); ++i)
			{
				if (ss_info[i].s == visSectionFirstComponent)
				{
					CString name = ss_info[i].name;
					name.Replace(L"{i}", FormatString(L"%d", 1 + s));
					name.Replace(L"{j}", FormatString(L"%d", 1 + r));

					if (StringIsLike(mask, name))
					{
						SRC src;

						src.name = name;
						src.s = visSectionFirstComponent + s;
						src.r = r;
						src.c = ss_info[i].c;

						result.push_back(src);
					}
				}
			}
		}
	}
}

void GetSimpleSectionCellNames(IVShapePtr shape, const CString& mask, std::vector<SRC>& result)
{
	for (size_t i = 0; i < countof(ss_info); ++i)
	{
		CString name = ss_info[i].name;

		if (StringIsLike(mask, name))
		{
			SRC src;

			src.name = name;
			src.s = ss_info[i].s;
			src.r = ss_info[i].r;
			src.c = ss_info[i].c;

			result.push_back(src);
		}
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
	GetVariableIndexedSectionCellNames(shape, visSectionCharacter, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionConnectionPts, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionLayer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionParagraph, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionReviewer, cell_name_mask, result);
	GetVariableIndexedSectionCellNames(shape, visSectionTextField, cell_name_mask, result);

	GetVariableGeometrySectionCellNames(shape, cell_name_mask, result);

	GetSimpleSectionCellNames(shape, cell_name_mask, result);
}


