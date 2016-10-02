using SeleniumAdvProject.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumAdvProject.DataObjects
{
    public class Panel
    {
        protected string _pageName;
        protected string _dataProfile;
        protected string _displayName;
        protected int _height;
        protected string _folder;
        public Panel() { }
        public Panel(Panel panel)
        {
            this._dataProfile = panel.DataProfile;
            this._displayName = panel.DisplayName;
            this._pageName = panel.PageName;
            this._height = panel.Height;
            this._folder = panel.Folder;
        }

        public Panel(string dataProfile, string displayName, string pageName, int height, string folder)
        {
            this._dataProfile = dataProfile;
            this._displayName = displayName;
            this._pageName = pageName;
            this._height = height;
            this._folder = folder;
        }

        public string DataProfile
        {
            get { return _dataProfile; }
            set { _dataProfile = value; }
        }
        public string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }
        public string PageName
        {
            get { return _pageName; }
            set { _pageName = value; }
        }
        public int Height
        {
            get { return _height; }
            set { _height = value; }
        }
        public string Folder
        {
            get { return _folder; }
            set { _folder = value; }
        }
       
    }
}
