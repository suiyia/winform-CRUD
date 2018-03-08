using StatisticalTools.po;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatisticalTools
{
    class BillPO
    {
        private int _id;  // 自增  1——。。。
        private String _billid;  // 单据编号： WPXSD171211
        private String _guestName;  //   客户名称
        private String _xinghao;  // 型号
        private String _zhonglei;  // 种类
        private String _color;  // 颜色
        private String _danwei;  // 单位
        private int _num;  // 数量
        private double _singlePrice;   // 单价
        private double _totalPrice;  // 总金额
        private String _kaipiaor; // 开票人
        private String _jinshour; // 经手人
        private String _picPath; // 上传的单据图片
        private String _kaipiaoDate; // 开票日期
        private String _address; // 厂址
        private String _otherText; // 备注

        public int Id
        {
            get
            {
                return _id;
            }

            set
            {
                _id = value;
            }
        }

        public string Billid
        {
            get
            {
                return _billid;
            }

            set
            {
                _billid = value;
            }
        }

        internal String GuestName
        {
            get
            {
                return _guestName;
            }

            set
            {
                _guestName = value;
            }
        }

        internal String Xinghao
        {
            get
            {
                return _xinghao;
            }

            set
            {
                _xinghao = value;
            }
        }

        internal String Zhonglei
        {
            get
            {
                return _zhonglei;
            }

            set
            {
                _zhonglei = value;
            }
        }

        internal String Color
        {
            get
            {
                return _color;
            }

            set
            {
                _color = value;
            }
        }

        public string Danwei
        {
            get
            {
                return _danwei;
            }

            set
            {
                _danwei = value;
            }
        }

        public int Num
        {
            get
            {
                return _num;
            }

            set
            {
                _num = value;
            }
        }

        public double SinglePrice
        {
            get
            {
                return _singlePrice;
            }

            set
            {
                _singlePrice = value;
            }
        }

        public double TotalPrice
        {
            get
            {
                return _totalPrice;
            }

            set
            {
                _totalPrice = value;
            }
        }

        public string Kaipiaor
        {
            get
            {
                return _kaipiaor;
            }

            set
            {
                _kaipiaor = value;
            }
        }

        public string Jinshour
        {
            get
            {
                return _jinshour;
            }

            set
            {
                _jinshour = value;
            }
        }

        public string PicPath
        {
            get
            {
                return _picPath;
            }

            set
            {
                _picPath = value;
            }
        }

        public String KaipiaoDate
        {
            get
            {
                return _kaipiaoDate;
            }

            set
            {
                _kaipiaoDate = value;
            }
        }

        public string Address
        {
            get
            {
                return _address;
            }

            set
            {
                _address = value;
            }
        }

        public string OtherText
        {
            get
            {
                return _otherText;
            }

            set
            {
                _otherText = value;
            }
        }
    }
}
