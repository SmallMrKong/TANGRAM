/********************************************************************************
*					Tangram Library - version 8.0								*
*********************************************************************************
* Copyright (C) 2002-2016 by Tangram Team.   All Rights Reserved.					*
*
* THIS SOURCE FILE IS THE PROPERTY OF TANGRAM TEAM AND IS NOT TO 
* BE RE-DISTRIBUTED BY ANY MEANS WHATSOEVER WITHOUT THE EXPRESSED 
* WRITTEN CONSENT OF TANGRAM TEAM.
*
* THIS SOURCE CODE CAN ONLY BE USED UNDER THE TERMS AND CONDITIONS 
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM 
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE 
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:sunhuizlz@yeah.net
*
********************************************************************************/
//#include "StdAfx.h"
#include "UccAPIEvent.h"

namespace OfficeCloudPlus
{
	namespace LyncPlus
	{
		namespace UccApiEvent
		{
			_ATL_FUNC_INFO OutgoingInvitation = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO IncomingInvitation = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO InvitationAccepted = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Renegotiate = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO SignalLevelChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO SourceChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO MediaStreamStateChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO PublishSessionMetrics = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO MediaRequest = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO UpdateChannels = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO MediaRequestCancelled = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO CategoryInstanceAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO CategoryInstanceRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO CategoryInstanceValueModified = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO EntityViewAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO EntityViewRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO PropertyUpdated = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO ScheduleConference = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO GetConference = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO GetConferenceList = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ModifyConference = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO DeleteConference = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO GetAvailableMcuList = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO ChannelAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ChannelRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO SetProperty = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO Enter = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Leave = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };


			//_ATL_FUNC_INFO UpdateChannels = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			_ATL_FUNC_INFO StateChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO AddEndpoint = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO RemoveEndpoint = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO NameChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ExternalUriChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO AddedToGroup = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO RemovedFromGroup = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ContactExtensionChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO MemberAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO MemberRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO ContainerAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ContainerRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO Enable = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Disable = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			//_ATL_FUNC_INFO NameChanged = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			//_ATL_FUNC_INFO ExternalUriChanged = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};

			_ATL_FUNC_INFO ContactAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ContactRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO GroupExtensionChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO SendUCMessage = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO InstantMessageReceived = { CC_STDCALL, VT_EMPTY, 2,VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Composing = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Idle = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			//_ATL_FUNC_INFO ChannelAdded = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			//_ATL_FUNC_INFO ChannelRemoved = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};

			_ATL_FUNC_INFO MediaDeviceChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO NegotiatedMediaChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO MediaDeviceAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO MediaDeviceRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO RecommendedMediaDeviceChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO SelectedMediaDeviceChanged = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO FindMediaConnectivityServers = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO CancelOperation = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO Shutdown = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };
			_ATL_FUNC_INFO IpAddrChange = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };

			_ATL_FUNC_INFO CategoryContextAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO CategoryContextRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO Publish = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO PublicationRequired = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO FindServer = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO TransferProgress = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Alternate = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO IncomingTransfer = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO AddParticipant = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };
			_ATL_FUNC_INFO RemoveParticipant = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };
			_ATL_FUNC_INFO Terminate = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };

			//_ATL_FUNC_INFO IncomingSession = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			//_ATL_FUNC_INFO OutgoingSession = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			_ATL_FUNC_INFO IncomingSession = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };
			_ATL_FUNC_INFO OutgoingSession = { CC_STDCALL,VT_EMPTY,2,VT_UNKNOWN,VT_UNKNOWN };

			_ATL_FUNC_INFO ParticipantAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO ParticipantRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO EndpointAdded = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO EndpointRemoved = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			//_ATL_FUNC_INFO StateChanged = {CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN};
			_ATL_FUNC_INFO AddParticipantEndpoint = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO RemoveParticipantEndpoint = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };

			_ATL_FUNC_INFO IncomingMessage = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO SendRequest = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };


			_ATL_FUNC_INFO Subscribe = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Unsubscribe = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
			_ATL_FUNC_INFO Query = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };


			_ATL_FUNC_INFO Execute = { CC_STDCALL, VT_EMPTY, 2, VT_UNKNOWN, VT_UNKNOWN };
		}
	}
}
