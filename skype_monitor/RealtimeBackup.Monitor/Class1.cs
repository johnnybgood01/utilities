using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RealtimeBackup.Monitor
{
    /// <summary>
    /// Provides access to a network share.
    /// </summary>
    public class NetworkShareAccesser : IDisposable
    {
        private string _remoteUncName;
        private string _remoteComputerName;

        public string RemoteComputerName
        {
            get
            {
                return this._remoteComputerName;
            }
            set
            {
                this._remoteComputerName = value;
                this._remoteUncName = @"\\" + this._remoteComputerName;
            }
        }

        public string UserName
        {
            get;
            set;
        }
        public string Password
        {
            get;
            set;
        }

        #region Consts

        private const int RESOURCE_CONNECTED = 0x00000001;
        private const int RESOURCE_GLOBALNET = 0x00000002;
        private const int RESOURCE_REMEMBERED = 0x00000003;

        private const int RESOURCETYPE_ANY = 0x00000000;
        private const int RESOURCETYPE_DISK = 0x00000001;
        private const int RESOURCETYPE_PRINT = 0x00000002;

        private const int RESOURCEDISPLAYTYPE_GENERIC = 0x00000000;
        private const int RESOURCEDISPLAYTYPE_DOMAIN = 0x00000001;
        private const int RESOURCEDISPLAYTYPE_SERVER = 0x00000002;
        private const int RESOURCEDISPLAYTYPE_SHARE = 0x00000003;
        private const int RESOURCEDISPLAYTYPE_FILE = 0x00000004;
        private const int RESOURCEDISPLAYTYPE_GROUP = 0x00000005;

        private const int RESOURCEUSAGE_CONNECTABLE = 0x00000001;
        private const int RESOURCEUSAGE_CONTAINER = 0x00000002;


        private const int CONNECT_INTERACTIVE = 0x00000008;
        private const int CONNECT_PROMPT = 0x00000010;
        private const int CONNECT_REDIRECT = 0x00000080;
        private const int CONNECT_UPDATE_PROFILE = 0x00000001;
        private const int CONNECT_COMMANDLINE = 0x00000800;
        private const int CONNECT_CMD_SAVECRED = 0x00001000;

        private const int CONNECT_LOCALDRIVE = 0x00000100;

        #endregion

        #region Errors

        private const int NO_ERROR = 0;

        private const int ERROR_ACCESS_DENIED = 5;
        private const int ERROR_ALREADY_ASSIGNED = 85;
        private const int ERROR_BAD_DEVICE = 1200;
        private const int ERROR_BAD_NET_NAME = 67;
        private const int ERROR_BAD_PROVIDER = 1204;
        private const int ERROR_CANCELLED = 1223;
        private const int ERROR_EXTENDED_ERROR = 1208;
        private const int ERROR_INVALID_ADDRESS = 487;
        private const int ERROR_INVALID_PARAMETER = 87;
        private const int ERROR_INVALID_PASSWORD = 1216;
        private const int ERROR_MORE_DATA = 234;
        private const int ERROR_NO_MORE_ITEMS = 259;
        private const int ERROR_NO_NET_OR_BAD_PATH = 1203;
        private const int ERROR_NO_NETWORK = 1222;

        private const int ERROR_BAD_PROFILE = 1206;
        private const int ERROR_CANNOT_OPEN_PROFILE = 1205;
        private const int ERROR_DEVICE_IN_USE = 2404;
        private const int ERROR_NOT_CONNECTED = 2250;
        private const int ERROR_OPEN_FILES = 2401;

        #endregion

        #region PInvoke Signatures

        [DllImport("Mpr.dll")]
        private static extern int WNetUseConnection(
            IntPtr hwndOwner,
            NETRESOURCE lpNetResource,
            string lpPassword,
            string lpUserID,
            int dwFlags,
            string lpAccessName,
            string lpBufferSize,
            string lpResult
            );

        [DllImport("Mpr.dll")]
        private static extern int WNetCancelConnection2(
            string lpName,
            int dwFlags,
            bool fForce
            );

        [StructLayout(LayoutKind.Sequential)]
        private class NETRESOURCE
        {
            public int dwScope = 0;
            public int dwType = 0;
            public int dwDisplayType = 0;
            public int dwUsage = 0;
            public string lpLocalName = "";
            public string lpRemoteName = "";
            public string lpComment = "";
            public string lpProvider = "";
        }

        #endregion

        /// <summary>
        /// Creates a NetworkShareAccesser for the given computer name. The user will be promted to enter credentials
        /// </summary>
        /// <param name="remoteComputerName"></param>
        /// <returns></returns>
        public static NetworkShareAccesser Access(string remoteComputerName)
        {
            return new NetworkShareAccesser(remoteComputerName);
        }

        /// <summary>
        /// Creates a NetworkShareAccesser for the given computer name using the given domain/computer name, username and password
        /// </summary>
        /// <param name="remoteComputerName"></param>
        /// <param name="domainOrComuterName"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        public static NetworkShareAccesser Access(string remoteComputerName, string domainOrComuterName, string userName, string password)
        {
            return new NetworkShareAccesser(remoteComputerName,
                                            domainOrComuterName + @"\" + userName,
                                            password);
        }

        /// <summary>
        /// Creates a NetworkShareAccesser for the given computer name using the given username (format: domainOrComputername\Username) and password
        /// </summary>
        /// <param name="remoteComputerName"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        public static NetworkShareAccesser Access(string remoteComputerName, string userName, string password)
        {
            return new NetworkShareAccesser(remoteComputerName,
                                            userName,
                                            password);
        }

        private NetworkShareAccesser(string remoteComputerName)
        {
            RemoteComputerName = remoteComputerName;

            this.ConnectToShare(this._remoteUncName, null, null, true);
        }

        private NetworkShareAccesser(string remoteComputerName, string userName, string password)
        {
            RemoteComputerName = remoteComputerName;
            UserName = userName;
            Password = password;

            this.ConnectToShare(this._remoteUncName, this.UserName, this.Password, false);
        }

        private void ConnectToShare(string remoteUnc, string username, string password, bool promptUser)
        {
            NETRESOURCE nr = new NETRESOURCE
            {
                dwType = RESOURCETYPE_DISK,
                lpRemoteName = remoteUnc
            };

            int result;
            if (promptUser)
            {
                result = WNetUseConnection(IntPtr.Zero, nr, "", "", CONNECT_INTERACTIVE | CONNECT_PROMPT, null, null, null);
            }
            else
            {
                result = WNetUseConnection(IntPtr.Zero, nr, password, username, 0, null, null, null);
            }

            if (result != NO_ERROR)
            {
                throw new Win32Exception(result);
            }
        }

        private void DisconnectFromShare(string remoteUnc)
        {
            int result = WNetCancelConnection2(remoteUnc, CONNECT_UPDATE_PROFILE, false);
            if (result != NO_ERROR)
            {
                throw new Win32Exception(result);
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <filterpriority>2</filterpriority>
        public void Dispose()
        {
            this.DisconnectFromShare(this._remoteUncName);
        }
    }

    public class FileMonitorConfiguration
    {
        public string SourceFilePath { get; set; }
        public string TargetFilePath { get; set; }
        public string FileName { get; set; }
       
    }

    public class RemoteFileMonitorConfiguration : FileMonitorConfiguration
    {
        public string ComputerName { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }

    public interface IChatMessage
    {
        string ConversationDisplayName { get; set; }

        void BuildCsv(StringBuilder builder);
    }

    public interface IChatConversation
    {
        int Id { get; set; }
        int IsPermanent { get; set; }
        int Type { get; set; }
        string LiveHost { get; set; }
        int LiveStartTimeStamp { get; set; }
        int LiveIsMuted { get; set; }
        string AlertString { get; set; }
        int IsBookMarked { get; set; }
        int IsBlocked { get; set; }
        string GivenDisplayName { get; set; }
        string DisplayName { get; set; }
        int LocalLiveStatus { get; set; }
        DateTime? InboxTimeStampAsDateTime { get; set; }
        int InboxTimeStamp { get; set; }
        int InboxMessageId { get; set; }
        int UnconsumedSuppressedMessages { get; set; }
        int UnconsumedNormalMessages { get; set; }
        int UnconsumedElevatedMessages { get; set; }
        int UnconsumedMessagesVoice { get; set; }
        int ActiveVmId { get; set; }
        int ContextHorizon { get; set; }
        int ConsumptionHorizon { get; set; }
        DateTime? LastActivityTimeStampAsDateTime { get; set; }
        int LastActivityTimeStamp { get; set; }
        int ActiveInvoiceMessage { get; set; }
        int SpawnedFromConvoId { get; set; }
        int PinnedOrder { get; set; }
        string Creator { get; set; }
        DateTime? CreationTimeStampAsDateTime { get; set; }
        int CreationTimeStamp { get; set; }
        int MyStatus { get; set; }
        int OptJoiningEnabled { get; set; }
        string OptAccessToken { get; set; }
        int OptEntryLevelRank { get; set; }
        int OptDiscloseHistory { get; set; }
        int OptHistoryLimitInDays { get; set; }
        int OptAdminOnlyActivities { get; set; }
        string PasswordHint { get; set; }
        string MetaName { get; set; }
        string MetaTopic { get; set; }
        string MetaGuidelines { get; set; }
        byte[] MetaPicture { get; set; }
        string Picture { get; set; }
        int IsP2PMigrated { get; set; }
        int PremiumVideoStatus { get; set; }
        int PremiumVideoIsGracePeriod { get; set; }
        string Guid { get; set; }
        string DialogPartner { get; set; }
        string MetaDescription { get; set; }
        string PremiumVideoSponsorList { get; set; }
        string McrCaller { get; set; }
        int ChatDbId { get; set; }
        int HistoryHorizon { get; set; }
        int HistorySyncState { get; set; }
        string ThreadVersion { get; set; }
        int ConsumptionHorizonSetAt { get; set; }
        string AltIdentity { get; set; }
        int InMigratedThreadSince { get; set; }
        int ExtPropProfileHeight { get; set; }
        int ExtPropChatWidth { get; set; }
        int ExtPropChatLeftMargin { get; set; }
        int ExtPropChatRightMargin { get; set; }
        int ExtPropEntryHeight { get; set; }
        int ExtPropWindowPosX { get; set; }
        int ExtPropWindowPosY { get; set; }
        int ExtPropWindowPosW { get; set; }
        int ExtPropWindowPosH { get; set; }
        int ExtPropWindowMaximized { get; set; }
        int ExtPropWindowDetached { get; set; }
        int ExtPropPinnedOrder { get; set; }
        int ExtPropNewInInbox { get; set; }
        int ExtPropTabOrder { get; set; }
        int ExtPropVideoLayout { get; set; }
        int ExtPropVideoChatHeight { get; set; }
        int ExtPropChatAvatar { get; set; }
        int ExtPropConsumptionTimeStamp { get; set; }
        int ExtPropFormVisible { get; set; }
        int ExtPropRecoveryMode { get; set; }
        int LastMessageId { get; set; }
        void BuildCsv(StringBuilder builder);

        /*
         id INTEGER NOT NULL PRIMARY KEY, 
         * is_permanent INTEGER, 
         * identity TEXT, 
         * type INTEGER, 
         * live_host TEXT, 
         * live_start_timestamp INTEGER, 
         * live_is_muted INTEGER, 
         * alert_string TEXT, 
         * is_bookmarked INTEGER, 
         * is_blocked INTEGER, 
         * given_displayname TEXT, 
         * displayname TEXT, 
         * local_livestatus INTEGER, 
         * inbox_timestamp INTEGER, 
         * inbox_message_id INTEGER, 
         * unconsumed_suppressed_messages INTEGER, 
         * unconsumed_normal_messages INTEGER, 
         * unconsumed_elevated_messages INTEGER, 
         * unconsumed_messages_voice INTEGER, 
         * active_vm_id INTEGER, 
         * context_horizon INTEGER, 
         * consumption_horizon INTEGER, 
         * last_activity_timestamp INTEGER, 
         * active_invoice_message INTEGER, 
         * spawned_from_convo_id INTEGER, 
         * pinned_order INTEGER, 
         * creator TEXT, 
         * creation_timestamp INTEGER, 
         * my_status INTEGER, 
         * opt_joining_enabled INTEGER, 
         * opt_access_token TEXT, 
         * opt_entry_level_rank INTEGER, 
         * opt_disclose_history INTEGER, 
         * opt_history_limit_in_days INTEGER, 
         * opt_admin_only_activities INTEGER, 
         * passwordhint TEXT, 
         * meta_name TEXT, 
         * meta_topic TEXT, 
         * meta_guidelines TEXT, 
         * meta_picture BLOB, 
         * picture TEXT, 
         * is_p2p_migrated INTEGER, 
         * premium_video_status INTEGER, 
         * premium_video_is_grace_period INTEGER, 
         * guid TEXT, 
         * dialog_partner TEXT, 
         * meta_description TEXT, 
         * premium_video_sponsor_list TEXT, 
         * mcr_caller TEXT, 
         * chat_dbid INTEGER, 
         * history_horizon INTEGER, 
         * history_sync_state TEXT, 
         * thread_version TEXT, 
         * consumption_horizon_set_at INTEGER, 
         * alt_identity TEXT, 
         * in_migrated_thread_since INTEGER, 
         * extprop_profile_height INTEGER, 
         * extprop_chat_width INTEGER, 
         * extprop_chat_left_margin INTEGER, 
         * extprop_chat_right_margin INTEGER, 
         * extprop_entry_height INTEGER, 
         * extprop_windowpos_x INTEGER, 
         * extprop_windowpos_y INTEGER, 
         * extprop_windowpos_w INTEGER, 
         * extprop_windowpos_h INTEGER, 
         * extprop_window_maximized INTEGER, 
         * extprop_window_detached INTEGER, 
         * extprop_pinned_order INTEGER, 
         * extprop_new_in_inbox INTEGER, 
         * extprop_tab_order INTEGER, 
         * extprop_video_layout INTEGER, 
         * extprop_video_chat_height INTEGER, 
         * extprop_chat_avatar INTEGER, 
         * extprop_consumption_timestamp INTEGER, 
         * extprop_form_visible INTEGER, 
         * extprop_recovery_mode INTEGER, 
         * last_message_id INTEGER
         */
    }

    public interface ISkypeDbAdapter
    { 
        void SetConnection(string pathToFile);
        IEnumerable<IChatConversation> GetConversations();
        IEnumerable<IChatMessage> GetChatMessages();

    }

    public class RemoteFileMonitor
    {
        private readonly RemoteFileMonitorConfiguration _configuration;
        protected ISkypeDbAdapter _skypeAdapter;
        protected IObservable<IChatMessage> ChatMessageStream;
        protected IObservable<IChatConversation> ChatRoomStream;
        private IDisposable
            _disposableNetworkConnection,
            _chatRoomSubscription,
            _chatMessageSubscription;
        protected string _targetFilePath;

        public RemoteFileMonitor(RemoteFileMonitorConfiguration configuration, ISkypeDbAdapter skypeAdapter)
        {
            _configuration = configuration;
            _skypeAdapter = skypeAdapter;
        }

        protected void StartInternal()
        {
            _disposableNetworkConnection = NetworkShareAccesser.Access(_configuration.ComputerName, null, _configuration.UserName, _configuration.Password);
            _skypeAdapter.SetConnection(string.Format("\\\\{0}\\{1}\\{2}", _configuration.ComputerName, _configuration.SourceFilePath, _configuration.FileName));
        }

        public void Start()
        {
                        StartInternal();
                        _targetFilePath = string.Format("{0}\\{1}.csv", _configuration.TargetFilePath, _configuration.FileName);
            _chatRoomSubscription =    ChatRoomStream.Subscribe(ProcessChatRoom);
            _chatMessageSubscription =   ChatMessageStream.Subscribe(ProcessChatMessage);
        }

        public void Stop()
        { 
             _chatRoomSubscription.Dispose();
             _chatMessageSubscription.Dispose();
             _disposableNetworkConnection.Dispose();

        }

        protected virtual void ProcessChatMessage(IChatMessage entry)
        {
            StringBuilder builder = new StringBuilder();
            entry.BuildCsv(builder);

            string builderResult = builder.ToString();

            if (string.IsNullOrWhiteSpace(builderResult))
                return;

            string conversationsFile = string.Format("{0}.conversations", _targetFilePath);

            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(conversationsFile, true))
            {
                writer.WriteLine(builderResult);
            }
        }

        protected virtual void ProcessChatRoom(IChatConversation conversationEntry)
        {
            StringBuilder builder = new StringBuilder();
            conversationEntry.BuildCsv(builder);

            string builderResult = builder.ToString();

            if (string.IsNullOrWhiteSpace(builderResult))
                return;

            string conversationsFile = string.Format("{0}.conversations", _targetFilePath);

            using(System.IO.StreamWriter writer = new System.IO.StreamWriter(conversationsFile, true))
            {
                writer.WriteLine(builderResult);
            }
        }
    }
}
