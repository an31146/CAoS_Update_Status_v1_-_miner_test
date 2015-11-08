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
    public class RevertDocStatus : IMANEXTLib.ICommand
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

        


        public RevertDocStatus()
        {
            // Intialise ICommand properties
            mName = "DocStatus Reverter";
            mTitle = "DocStatus Reverter";
            mAccelerator = 0;
            mType = IMANEXTLib.CommandType.nrStandardCommand;
            mStatus = (int)IMANEXTLib.CommandStatus.nrActiveCommand;
            mMenuText = "DocStatus Reverter";
            mHelpText = "DocStatus Reverter";
        }

        public void Initialize(IMANEXTLib.ContextItems Context)
        {
            mContext = Context;
            mStatus = (int)IMANEXTLib.CommandStatus.nrActiveCommand;
        }

        public void Update()
        {
            
        }

        public void Execute()
        {
            MessageBox.Show("Revert Status invoked");
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
