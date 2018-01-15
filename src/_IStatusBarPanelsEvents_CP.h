//////////////////////////////////////////////////////////////////////
/// \class Proxy_IStatusBarPanelsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IStatusBarPanelsEvents interface</em>
///
/// \if UNICODE
///   \sa StatusBarPanels, StatBarLibU::_IStatusBarPanelsEvents
/// \else
///   \sa StatusBarPanels, StatBarLibA::_IStatusBarPanelsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IStatusBarPanelsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IStatusBarPanelsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IStatusBarPanelsEvents