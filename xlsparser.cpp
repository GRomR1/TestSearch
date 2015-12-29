#include "xlsparser.h"

XlsParser::XlsParser(QObject *parent) :
    QObject(parent)
  , _error(false)
  , _isOneColumn(false)
{
    qDebug() << "XlsParser constructor" << this->thread()->currentThreadId();
}

XlsParser::~XlsParser()
{
    qDebug() << "XlsParser destructor" << this->thread()->currentThreadId();
}

void XlsParser::process()
{

}

void XlsParser::onReadHead(const QString sheet,
                           MapAddressElementPosition head)
{
//    qDebug() << "XlsParser onReadHead"
//             << sheet << head.values()
//             << this->thread()->currentThreadId();
    _mapHead.insert(sheet, head);
}

void XlsParser::onIsOneColumn(bool b)
{
    _isOneColumn=b;
}

void XlsParser::onReadRow(const QString &sheet,
                          const int &rowNumber,
                          const QStringList &row)
{
//    qDebug() << "XlsParser onReadRow" << sheet << rowNumber
//             /*<< row */<< this->thread()->currentThreadId() ;

    Address a;
//    a.setRawAddress(row);
    //парсинг строки начат
    QString str;
    QStringList resList;
    QString ptrn;

    //работаем с CITY
    if(_mapHead[sheet].contains(CITY1))
    {
        str = row.at(_mapHead[sheet].value(CITY1));
        str = str.trimmed();
        str = str.toLower();

        //с городом
        //учитывает только г. но запись с двух сторон возможна
//        ptrn = "((?:(?:^)|(?:[.,;()])\\s*)"
//               "(?:(?:(?:(город|гор\\.*|г\\.*)\\s+)*([\\w\\d\\s-]+))|"
//               "(?:([\\w\\d\\s-]+)(?:\\s+(город|гор\\.*|г\\.*))*)\\s*)"
//               "(?:(?:[,;()])|(?:$))";

        //запись название города только после типа города
        ptrn = "(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|[сдп]\\.*|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)\\s+([\\w\\d\\s-]+)(?:([,;]|$)|[()])";

        //возможна запись имени города без указания типа
        //забирает первое слово до разделителя!
//        ptrn = "(?:(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|[сдп]\\.*|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)\\s+)*([\\w\\d\\s-]+)(?:([,;]|$)|[()])";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
            str.remove(resList.at(1)+resList.at(2)+resList.at(3));
            a.setTypeOfCity1(resList.at(1));
            a.setCity1(resList.at(2));
        }
    }

    //работаем с STREET
    if(_mapHead[sheet].contains(STREET))
    {
        str = row.at(_mapHead[sheet].value(STREET));
        str = str.trimmed();
        str = str.toLower();

        //работаем с федеральным субъектом
//        ptrn = "((?:^|[.,;()]\\s*)"
//               "(?:(?:([\\w\\d\\s-]+)\\s+(?:республика|респ\\.*|область|обл\\.*|край|ао|aобл\\.*))|"
//               "(?:(?республика|респ\\.*|ао|aобл\\.*)\\s+([\\w\\d\\s-]+))))"
//               "(?:[.,;()]|$)";
//        ptrn = "((?:^|[.,;()]\\s*)"
//               "(?:(?:([\\w\\d\\s-]+)\\s+(?:республика|респ\\.*|область|обл\\.*|край|ао|aобл\\.*))|"
//               "(?:(?:республика|респ\\.*|ао|aобл\\.*)\\s+([\\w\\d\\s-]+))))"
//               "(?:[.,;()]|$)";
        ptrn="((?:^|[.,;()]\\s*)(?:(?:([\\w\\d\\s-]+)\\s+(республика|респ\\.*|область|обл\\.*|край|ао|aобл\\.*))|(?:(республика|респ\\.*|ао|aобл\\.*)\\s+([\\w\\d\\s-]+))))(?:[.,;()]|$)";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
            str.remove(resList.at(1));
            if(!resList.at(2).isEmpty())
                a.setFsubj(resList.at(2));
            else if(!resList.at(5).isEmpty())
                a.setFsubj(resList.at(5));
            if(!resList.at(3).isEmpty())
                a.setTypeOfFSubj(resList.at(3));
            else if(!resList.at(4).isEmpty())
                a.setTypeOfFSubj(resList.at(4));
        }
        //с районом
        ptrn = "((?:^|[.,;()]\\s*)"
               "(?:(?:([\\w\\d\\s-]+)\\s+(?:район|р-н))|"
               "(?:(?:район|р-н)\\s+([\\w\\d\\s-]+))))"
               "(?:[.,;()]|$)";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
            str.remove(resList.at(1));
            if(!resList.at(2).isEmpty())
                a.setDistrict(resList.at(2));
            else if(!resList.at(3).isEmpty())
                a.setDistrict(resList.at(3));
        }
        //с городом
//        ptrn = "((?:^|[.,;()]\\s*)"
//               "(?:(?:(?:город|гор\\.*|г\\.*)\\s+([\\w\\d\\s-]+))|"
//               "(?:([\\w\\d\\s-]+)\\s+(?:город|гор\\.*|г\\.*))))"
//               "(?:[.,;()]|$)";
//        ptrn = "(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|[сдп]\\.*|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)(\\s+[\\w\\d\\s.-]+)(?:([.,;]|$)|[()])";
        //учитывает пропуск вначале и дом от деревни (по следующей цифре за "д.")
        ptrn="\\b(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|(?:[сдп]\\.*(?!\\s*\\d+))|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)(\\s+[\\w\\d\\s.-]+)(?:([.,;]|$)|[()])";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
//            str.remove(resList.at(1));
//            if(!resList.at(2).isEmpty())
//                a.setCity(resList.at(2));
//            else if(!resList.at(3).isEmpty())
//                a.setCity(resList.at(3));

            str.remove(resList.at(1)+resList.at(2)+resList.at(3));
            a.setTypeOfCity1(resList.at(1));
            a.setCity1(resList.at(2));
        }

//        ptrn = "(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|[сдп]\\.*|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)(\\s+[\\w\\d\\s.-]+)(?:([.,;]|$)|[()])";
        //учитывает пропуск вначале и дом от деревни (по следующей цифре за "д.")
        ptrn="\\b(село|деревня|дер\\.*|поселок|пос\\.*|станция|ст\\.*|(?:[сдп]\\.*(?!\\s*\\d+))|город|гор\\.*|г\\.*|ж\\/д_рзд\\.|рп\\.|ж\\/д_ст\\.|высел\\.|х\\.*|казарма\\.*|м\\.|жилрайон\\.*|нп\\.*|снт\\.*|дп\\.*|снт\\.*|пст\\.|п\\/ст\\.|пгт\\.|п\\/гт\\.|тер\\.*|массив\\.*)(\\s+[\\w\\d\\s.-]+)(?:([.,;]|$)|[()])";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
//            str.remove(resList.at(1));
//            if(!resList.at(2).isEmpty())
//                a.setCity(resList.at(2));
//            else if(!resList.at(3).isEmpty())
//                a.setCity(resList.at(3));

            str.remove(resList.at(1)+resList.at(2)+resList.at(3));
            a.setTypeOfCity2(resList.at(1));
            a.setCity2(resList.at(2));
        }

        //с улицей, пр-том, и пр.
        ptrn = "((?:[.,;()]*\\s*)"
               "(?:(?:(улица|ул\\.*|шоссе|ш\\.*|аллея|ал\\.*|бульвар|б-р|линия|лин\\.*|набережная|наб\\.*|парк|переулок|пер\\.*|площадь|пл\\.*|проспект|пр-кт|пр\\.*|сад|сквер|строение|стр\\.*|участок|уч-к|квартал|берег|кв-л|километр|км|кольцо|проезд|переезд|въезд|заезд|дорога|дор\\.*|платформа|платф\\.*|площадка|пл-ка\\.*|полустанок|полуст\\.*|проулок|просека|тракт|тупик|туп\\.*|просёлок)\\s+([\\w\\d\\s-.]+))|"
               "(?:([\\w\\d\\s-.]+)\\s+(улица|ул\\.*|шоссе|ш\\.*|аллея|ал\\.*|бульвар|б-р|линия|лин\\.*|набережная|наб\\.*|парк|переулок|пер\\.*|площадь|пл\\.*|проспект|пр-кт|пр\\.*|сад|сквер|строение|стр\\.*|участок|уч-к|квартал|берег|кв-л|километр|км|кольцо|проезд|переезд|въезд|заезд|дорога|дор\\.*|платформа|платф\\.*|площадка|пл-ка\\.*|полустанок|полуст\\.*|проулок|просека|тракт|тупик|туп\\.*|просёлок)))\\s*)"
               "(?:([,;]|$)|[()])";
        if(parseObject(str,
                       resList,
                       ptrn))
        {
            str.remove(resList.at(1));
            if(!resList.at(2).isEmpty())
                a.setTypeOfStreet(resList.at(2));
            else if(!resList.at(5).isEmpty())
                a.setTypeOfStreet(resList.at(5));
            if(!resList.at(3).isEmpty())
                a.setStreet(resList.at(3));
            else if(!resList.at(4).isEmpty())
                a.setStreet(resList.at(4));
        }

        //работа со скобками
        int n1=str.indexOf('(');
        if (n1>0 && (str.indexOf(')',n1)>0))
        {
            QString additional=a.getAdditional()+" ";
            int n2=str.indexOf(')', n1);
            int n3=n2-n1;
            additional+=str.mid(n1+1, n3-1);
            str.remove(n1, n3+1);
            a.setAdditional(additional);
        }

        //если в этом же столбце находится номер дома и корпус
        // TODO: check it!
//        if(_isOneColumn)
//        {
//            qDebug() << "isOneColumn";
            //дом, корпус и литера
            ptrn = "(?:[.,;]*\\s*)(д\\.*|дом\\.*|ая\\.*|а\\\\я\\.*|а/я\\.*)\\s*([0-9Х][-\\\\/0-9]*)([а-я]*)\\s*,*\\s*(?:(к\\.*|корп\\.*)\\s*(\\d*)-*(\\w*))*\\s*,*\\s*(?:(л\\.*|лит\\.*|литера\\.*)\\s*(\\d*)-*(\\w*))*\\s*(?:(?:кв\\.*)|([а-я]*))";
            if(parseObject(str,
                           resList,
                           ptrn))
            {
                str.remove(resList.at(0));
                a.setBuild(resList.at(2));
                QString k;
                if(!resList.at(5).isEmpty())
                    k=resList.at(5);
                else
                    k=resList.at(8);
                QString l;
                l=resList.at(3)+resList.at(10)+resList.at(6)+
                        resList.at(10)+resList.at(9);
                a.setKorp(k);
                a.setLitera(l);
//            }
        }
    } // конец работы с STR
/*
    //если номер дома и корпус находятся в другом столбце
    if(!_isOneColumn && a.getBuild().isEmpty()
                          && a.getKorp().isEmpty()
                            && a.getLitera().isEmpty())
    {
//        qDebug() << "! isOneColumn";
        //работаем с B
        if(_mapHead[sheet].contains(BUILD))
        {
            QString str = row.at(_mapHead[sheet].value(BUILD));
            str = str.trimmed();
            str = str.toLower();
            //дом
            ptrn = "(?:[.,;]*\\s*)(д\\.*|дом\\.*|ая\\.*|а\\\\я\\.*|а/я\\.*)*\\s*([0-9Х][-\\\\/0-9]*)([а-я]*)\\s*,*\\s*(?:(к\\.*|корп\\.*)\\s*(\\d*)-*(\\w*))*\\s*,*\\s*(?:(л\\.*|лит\\.*|литера\\.*)\\s*(\\d*)-*(\\w*))*\\s*([а-я]*)";
            if(parseObject(str,
                           resList,
                           ptrn))
            {
                str.remove(resList.at(0));
                a.setBuild(resList.at(2));
                QString k;
                if(!resList.at(5).isEmpty())
                    k=resList.at(5);
                else
                    k=resList.at(8);
                QString l;
                l=resList.at(3)+resList.at(10)+resList.at(6)+
                        resList.at(10)+resList.at(9);
                a.setKorp(k);
                a.setLitera(l);
            }
        }//end work with B

        //работаем с K
        if(_mapHead[sheet].contains(KORP))
        {
            QString str = row.at(_mapHead[sheet].value(KORP));
            str = str.trimmed();
            str = str.toLower();
            //корпус и литера
            ptrn = "\\s*(?:(к\\.*|корп\\.*)*\\s*(\\d*)-*(\\w*))*\\s*,*\\s*(?:(л\\.*|лит\\.*|литера\\.*)\\s*(\\d*)-*(\\w*))*\\s*([а-я]*)";
            if(parseObject(str,
                           resList,
                           ptrn))
            {
                str.remove(resList.at(0));
                QString k;
                if(!resList.at(2).isEmpty())
                    k=resList.at(2);
                else
                    k=resList.at(5);
                QString l;
                l=resList.at(3)+resList.at(6);
                a.setKorp(k);
                a.setLitera(l);
            }
        }//end work with K
    }//если номер дома и корпус находятся в другом столбце
    */
    //парсинг строки окончен

//    qDebug().noquote() << " XlsParser::" << str << a.toString(PARSED);
//    assert(0);
    emit rowParsed(sheet, rowNumber, a); //парсинг строки окончен
//    return a;
}


bool XlsParser::parseObject(QString &str, QString &result,
                            QString regPattern, int regCap)
{
    if(str.isEmpty())
        return false;
    QRegExp rx(regPattern);
    int pos=rx.indexIn(str);
    if(pos!=-1)
    {
        str.remove(rx.cap(0)); //удаляем все что нашли
        result = rx.cap(regCap).trimmed();
        return true;
    }
    else
        return false;
}

bool XlsParser::parseObject(QString &str, QStringList &result,
                            QString regPattern)
{
    if(str.isEmpty())
        return false;
    QRegExp rx(regPattern);
    int pos=rx.indexIn(str);
    if(pos!=-1)
    {
//        str.remove(rx.cap(0)); //удаляем все что нашли
        result = rx.capturedTexts();
        return true;
    }
    else
        return false;
}
