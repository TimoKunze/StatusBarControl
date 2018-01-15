//////////////////////////////////////////////////////////////////////
/// \class Proxy_IStatusBarPanelEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IStatusBarPanelEvents interface</em>
///
/// \if UNICODE
///   \sa StatusBarPanel, StatBarLibU::_IStatusBarPanelEvents
/// \else
///   \sa StatusBarPanel, StatBarLibA::_IStatusBarPanelEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IStatusBarPanelEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IStatusBarPanelEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IStatusBarPanelEvents