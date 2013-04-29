#if !defined(__GRID_EXCEL_EXPORT_H__)
#define __GRID_EXCEL_EXPORT_H__

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include "GridCtrl.h"

namespace utils {
namespace grid {

void OpenGridInExcel(CGridCtrl* grid);

}}

#endif //__GRID_EXCEL_EXPORT_H__