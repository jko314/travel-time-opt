package com.leidos.tamp.type;

public class Global {
    // TODO: Option Explicit ... Warning!!! not translated
//
//
//    public final long decodeFromValueList(String  strValue, String  strValueList) {
//        long lngStartOffset;
//        long lngEndOffset;
//        long decodeFromValueListVal;
//
//        lngEndOffset = (strValueList.indexOf(("," + (strValue + ";"))) + 1);
//        if ((lngEndOffset == 0)) {
//            decodeFromValueListVal = -1;
//        }
//        else {
//            lngStartOffset = (InStrRev(strValueList, ";", lngEndOffset) + 1);
//            decodeFromValueListVal = Long.parseLong(strValueList.substring((lngStartOffset - 1), (lngEndOffset - lngStartOffset)));
//        }
//
//        return decodeFromValueList;
//    }
//
//    public final String  encodeFromValueList(long lngValue, String  strValueList) {
//        long lngStartOffset;
//        long lngEndOffset;
//        String  strValue;
//        strValue = (lngValue.ToString() + ",");
//        if ((strValueList.Substring(0, strValue.Length) == strValue)) {
//            lngStartOffset = (1 + strValue.Length);
//        }
//        else {
//            lngStartOffset = (strValueList.IndexOf((";" + strValue)) + 1);
//            if ((lngStartOffset == 0)) {
//                // TODO: Exit Function: Warning!!! Need to return the value
//            }
//
//            return;
//            lngStartOffset = (lngStartOffset
//                    + (strValue.Length + 1));
//        }
//
//        lngEndOffset = (strValueList.IndexOf(";", (lngStartOffset - 1)) + 1);
//        return strValueList.Substring((lngStartOffset - 1), (lngEndOffset - lngStartOffset));
//    }

}
