using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Configuration;
using System.Net;
using System.Xml;
using System.Data.SqlClient;
using System.Reflection;
using System.IO;
using IMANEXTLib;
using IManage;
using System.Diagnostics;

namespace UpdateStatus
{
    public class ParentRefresh : IMANEXTLib.ICommand
    {
        // ICommand properties
        private int mAccelerator;
        private object mBitmap;
        private IMANEXTLib.ContextItems mContext;
        private string mHelpFile;
        private int mHelpID;
        private string mHelpText;
        private string mMenuText;
        private string mName;
        private int mOptions;
        private int mStatus;
        private IMANEXTLib.Commands mSubCommands;
        private string mTitle;
        private IMANEXTLib.CommandType mType;
        

        
        public ParentRefresh()
        {
            // Intialise ICommand properties
            mName = "DocStatus Refresh";
            mTitle = "DocStatus Refresh";
            mAccelerator = 0;
            mType = IMANEXTLib.CommandType.nrStandardCommand;
            mStatus = (int)IMANEXTLib.CommandStatus.nrActiveCommand;
            mMenuText = "DocStatus Refresh";
            mHelpText = "DocStatus Refresh";
        }

        public void Initialize(IMANEXTLib.ContextItems Context)
        {

            mContext = Context;
            mStatus = (int)IMANEXTLib.CommandStatus.nrActiveCommand;
        }

        public void Update()
        {
        }

        public void DeleteFolder(IManFolder fold)
        {
            NRTDMS dms = new NRTDMS();
            dms.Sessions.Add("172.16.31.139");
            IManSession sess = (IManSession)dms.Sessions.Item(1);
            sess.Login("wsadmin", "mhdocs_");
            IManDatabase db = sess.Databases.ItemByName("WorkSite83");
            IManFolder work = db.GetFolder(fold.Parent.FolderID);
            IManFolder fldr1 = db.GetFolder(fold.FolderID);
            IManDocumentFolders docfolds = (IManDocumentFolders)work.SubFolders;

            Debug.WriteLine(work.Name + " - " + work.SubFolders.Count.ToString());
            Debug.WriteLine("IsOpAllowed: " + work.SubFolders.IsOperationAllowed(imFoldersOp.imRemoveFoldersOp).ToString());
            Debug.WriteLine("contains: " + docfolds.Contains(fldr1).ToString());

            if (docfolds.Contains(fldr1))
            {
                docfolds.RemoveByObject(fldr1);
                Debug.WriteLine(fold.Name + " ... deleted.");
            }
            sess.Logout();
        }

        public void Execute()
        {
            IManSession sess = null;
            Object[] obs = mContext.Item("SelectedNRTSessions") as Object[];
            foreach (Object obj2 in obs)
            {
                sess = (IManSession)obj2;
                break;
            }

            MessageBox.Show("ParentRefresh invoked");
            mContext.Add("IManExt.Refresh", true);
            mContext.Add("RefreshSubFolders", true);
            mContext.Add("RefreshAllFolders", true);
            
            IManage.IManFolder fldr = mContext.Item("SelectedFolderObject") as IManage.IManFolder;
            Debug.WriteLine(fldr.Name + " ... " + fldr.ObjectID.ToString());

            DeleteFolder(fldr);

            IManage.IManFolder parent = fldr.Parent as IManage.IManFolder;
            IManFolder parent1 = sess.DMS.GetObjectBySID(fldr.ObjectID) as IManFolder;
            Debug.WriteLine(parent1.Name + " ... " + parent1.ObjectID.ToString());

            if (parent != null)
            {
                Debug.WriteLine("Refresh parent container");
                //parent1.Refresh();
                parent1.SubFolders.Refresh();
            }
        }


        #region donotmodify
        public string HelpText
        {
            get
            {
                return mHelpText;
            }
            set
            {
                mHelpText = value;
            }
        }

        public string Title
        {
            get
            {
                return mTitle;
            }
            set
            {
                mTitle = value;
            }
        }

        public string MenuText
        {
            get
            {
                return mMenuText;
            }
            set
            {
                mMenuText = value;
            }
        }

        public object Bitmap
        {
            get
            {
                return mBitmap;
            }
            set
            {
                mBitmap = value;
            }
        }

        public int HelpID
        {
            get
            {
                return mHelpID;
            }
            set
            {
                mHelpID = value;
            }
        }

        public IMANEXTLib.ContextItems Context
        {
            get
            {
                return mContext;
            }
        }

        public IMANEXTLib.Commands SubCommands
        {
            get
            {
                return mSubCommands;
            }
            set
            {
                mSubCommands = value;
            }
        }

        public string Name
        {
            get
            {
                return mName;
            }
            set
            {
                mName = value;
            }
        }

        public string HelpFile
        {
            get
            {
                return mHelpFile;
            }
            set
            {
                mHelpFile = value;
            }
        }

        public IMANEXTLib.CommandType Type
        {
            get
            {
                return mType;
            }
            set
            {
                mType = value;
            }
        }

        public int Accelerator
        {
            get
            {
                return mAccelerator;
            }
            set
            {
                mAccelerator = value;
            }
        }

        public int Options
        {
            get
            {
                return mOptions;
            }
            set
            {
                mOptions = 0;
            }
        }

        public int Status
        {
            get
            {
                return mStatus;
            }
            set
            {
                mStatus = value;
            }
        }
        #endregion

    }
}
