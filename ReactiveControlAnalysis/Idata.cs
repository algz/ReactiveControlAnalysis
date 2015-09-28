using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReactiveControlAnalysis
{
    public class Idata
    {
        public Idata(int id, string name, string value1, string error1, string value2, string error2, string value3, string error3, string errorMax)
        {
            if (error1 != "0" && errorMax != "" && value1 != "")
                this._error1 = (Convert.ToDouble(value1) * Convert.ToDouble(errorMax.Remove (errorMax .Length -1)) * 0.01).ToString();
            else
                this._error1 = error1;
            this._id = id;
            this._value1 = value1;
            this._name = name;

            if (value2 != "0")
                this._value2 = this._value1;
            else
                this._value2 = value2;
            if (error2 != "0")
                this._error2 = this._error1;
            else
                this._error2 = error2;

            if (value3 != "0")
            {
                if (value3 == "空")
                    this._value3 = "";
                else
                    this._value3 = this._value1;
            }
            else
                this._value3 = value3;
            if (error3 != "0")
            {
                if (error3 == "空")
                    this._error3 = "";
                else
                    this._error3 = this._error1;
            }
            else
                this._error3 = error3;

            this._errorMax = errorMax;
        }
        int _id;

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        string _name, _value1, _error1, _value2, _error2, _value3, _error3, _errorMax;

        public string ErrorMax
        {
            get
            {
                return _errorMax;
            }
            set
            {
                _errorMax = value;
                this._error1 = (Convert.ToDouble(this._value1) * Convert.ToDouble(_errorMax.Remove(_errorMax.Length - 1)) * 0.01).ToString();
           
            }
        }

        public string Error3
        {
            get { return _error3; }
            set { _error3 = value; }
        }

        public string Value3
        {
            get { return _value3; }
            set { _value3 = value; }
        }

        public string Error2
        {
            get { return _error2; }
            set { _error2 = value; }
        }

        public string Value2
        {
            get { return _value2; }
            set { _value2 = value; }
        }

        public string Error1
        {
            get { return _error1; }
            set
            {
                _error1 = value;
            }
        }

        public string Value1
        {
            get { return _value1; }
            set { _value1 = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }




    }
}
