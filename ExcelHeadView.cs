using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Panduo.NPOIDemo
{
    public class ExcelHeadView
    {
        /// <summary>
        /// ����
        /// </summary>
        public string CellName { get; set; }
        /// <summary>
        /// ���б�ͷ���Ƿ��ǵ�һ������
        /// </summary>
        public bool IsFirstSonCell { get; set; }
        /// <summary>
        /// ���б�ͷ�����е�����
        /// </summary>
        public int Span { get; set; }
        /// <summary>
        /// ���б�ͷ������
        /// </summary>
        public string FatherCellName { get; set; }
        /// <summary>
        /// ������
        /// </summary>
        public string TableName { get; set; }
        /// <summary>
        /// �����ڱ������λ��
        /// </summary>
        public int index { get; set; }

    }
}