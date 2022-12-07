<?php

namespace lib\usertype;

use Bitrix\Highloadblock\HighloadBlockTable;

/**
 * Реализация свойства «Конфигурация»
 * Class CUserTypeTimesheet
 * @package lib\usertype
 */
class CUserTypeConfiguration
{
    /**
     * Метод возвращает массив описания собственного типа свойств
     * @return array
     */
    public function GetUserTypeDescription()
    {
        return array(
            'USER_TYPE_ID' => 'user_configuration',
            'USER_TYPE' => 'Configuration',
            'CLASS_NAME' => __CLASS__,
            'DESCRIPTION' => 'Группы товаров, с выбором товара для каждой группы',
            'PROPERTY_TYPE' => 'S',
            'ConvertToDB' => [__CLASS__, 'ConvertToDB'],
            'ConvertFromDB' => [__CLASS__, 'ConvertFromDB'],
            'GetPropertyFieldHtml' => [__CLASS__, 'GetPropertyFieldHtml'],
        );
    }

    /**
     * Конвертация данных перед сохранением в БД
     * @param $arProperty
     * @param $value
     * @return mixed
     */
    public static function ConvertToDB($arProperty, &$value)
    {
        $value['VALUE']['PRODUCTS'] = is_array($value['VALUE']['PRODUCTS']) ? array_unique(array_diff($value['VALUE']['PRODUCTS'], [''])) : $value['VALUE']['PRODUCTS'];
        $value['VALUE']['PRODUCTS'] = array_values($value['VALUE']['PRODUCTS']);

        if ($value['VALUE']['GROUP'] != '' && $value['VALUE']['PRODUCTS'] != '' && !empty($value['VALUE']['PRODUCTS'])) {
            try {
                $value['VALUE'] = base64_encode(serialize($value['VALUE']));
            } catch (Bitrix\Main\ObjectException $exception) {
                echo $exception->getMessage();
            }
        } else {
            $value['VALUE'] = '';
        }
        return $value;
    }

    /**
     * Конвертируем данные при извлечении из БД
     * @param $arProperty
     * @param $value
     * @param string $format
     * @return mixed
     */
    public static function ConvertFromDB($arProperty, $value, $format = '')
    {
        if ($value['VALUE'] != '') {
            try {
                $value['VALUE'] = base64_decode($value['VALUE']);
            } catch (Bitrix\Main\ObjectException $exception) {
                echo $exception->getMessage();
            }
        }
        return $value;
    }

    /**
     * Представление формы редактирования значения
     * @param $arUserField
     * @param $arHtmlControl
     */
    public static function GetPropertyFieldHtml($arProperty, $value, $arHtmlControl)
    {
        $index = 0;

        $hlData = HighloadBlockTable::getById(3)->fetch();
        $hlEntity = HighloadBlockTable::compileEntity($hlData);
        $entityDataClass = $hlEntity->getDataClass();
        $rsData = $entityDataClass::getList(array(
            'order' => array('UF_NAME' => 'ASC'),
            'select' => array('*'),
            'filter' => array('!UF_NAME' => false)
        ));

        $groups = [];
        while ($arData = $rsData->Fetch()) {
            $groups[$arData['ID']] = $arData['UF_NAME'];
        }

        $permitted_chars = '0123456789abcdefghijklmnopqrstuvwxyz';

        $itemId = 'row_' . substr(str_shuffle($permitted_chars), 0, 10);
        $fieldName = htmlspecialcharsbx($arHtmlControl['VALUE']);
        $arValue = unserialize(htmlspecialcharsback($value['VALUE']), [stdClass::class]);

        $productsResult = \Bitrix\Iblock\ElementTable::getList([
            'select' => ['ID', 'NAME'],
            'filter' => ['ID' => $arValue['PRODUCTS']],
        ]);

        $productsNames = [];
        while ($prod = $productsResult->fetch()) {
            $productsNames[$prod['ID']] = $prod['NAME'];
        }

        $html = '<div class="property_row conf-row" id="' . $itemId . '">';
        $html .= '<div class="configuration">';

        $select = '<select class="group conf-select" name="'
            . $fieldName
            . '[GROUP]" onchange="let blockConf = this.parentNode.parentNode, id = blockConf.id;'
            . 'blockConf.id ='
            . "'row_'"
            . '+(Math.random()*1e32).toString(36).substring(0, 10);'
            . 'let delBtn = this.parentNode.querySelector(' . "'.conf-btn-del'" . ');'
            . 'let newFunction = '
            . "'document.getElementById('"
            . "+'\''"
            . '+ blockConf.id +'
            . "'\''+"
            . "').parentNode.parentNode.remove()';"
            . 'delBtn.setAttribute(\'onclick\', newFunction);'
            . '" > ';
        foreach ($groups as $key => $group) {
            if ($arValue['GROUP'] == $key) {
                $select .= '<option value = "' . $key . '" selected = "selected" > ' . $group . '</option>';
            } else {
                $select .= '<option value = "' . $key . '" > ' . $group . '</option> ';
            }
        }

        $products = ($arValue['PRODUCTS']) ?: '';
        $maxCount = ($arValue['MAX_COUNT']) ?: '';
        $recCount = ($arValue['REC_COUNT']) ?: '';

        $select .= ' </select> ';

        $html .= $select;
        $html .= ' &nbsp; Max кол - во:&nbsp;<input style = "width: 50px;" type = "number" min = 0 name = "' . $fieldName . '[MAX_COUNT]" size = "5" value = "' . $maxCount . '" > ';
        $html .= ' &nbsp; Рекомендуемое кол - во:&nbsp;<input style = "width: 50px;" type = "number" min = 0 name = "' . $fieldName . '[REC_COUNT]" size = "5" value = "' . $recCount . '" > ';

        $html .= ' &nbsp;&nbsp;<input class="conf-btn-del" type = "button" style = "height: auto;" value = "x" title = "Удалить" onclick = "document.getElementById(\'' . $itemId . '\').parentNode.parentNode.remove()" />';

        $html .= '<br/><br/> Товары:';
        $html .= '<table class="conf-tbl" id="dv' . md5($fieldName) . '" data-field-name="'. $fieldName .  '">';
        if (empty($products)) {
            $html .= '<tr><td>&nbsp; <input type="text" name="' . $fieldName . '[PRODUCTS][0]" size ="5" id="' . $fieldName . '[PRODUCTS][0]">';
            $html .= '&nbsp; <input type="button" value="..."
            onClick="jsUtils.OpenWindow(\'/bitrix/admin/iblock_element_search.php?lang=' . LANGUAGE_ID . '&amp;IBLOCK_ID=5&amp;n=' . $fieldName . '[PRODUCTS][0]' . '&amp;iblockfix=y' . '\', 900, 700);">';
            $html .= '<span id="sp_' . $fieldName . '[PRODUCTS][0]' . '"></span></td></tr>';
            $index++;
        } else {
            for ($i = 0; $i < count($products) + 1; ++$i) {
                $html .= '<tr><td>&nbsp; <input type="text" name="' . $fieldName . '[PRODUCTS][' . $i . ']" size ="5" value="' . $products[$i] . '" id="' . $fieldName . '[PRODUCTS][' . $i . ']">';
                $html .= '&nbsp; <input type="button" value="..."
                onClick="jsUtils.OpenWindow(\'/bitrix/admin/iblock_element_search.php?lang=' . LANGUAGE_ID . '&amp;IBLOCK_ID=5&amp;k=n0&amp;n=' . $fieldName . '[PRODUCTS][' . $i . ']' . '&amp;iblockfix=y' . '\', 900, 700);">';
                $html .= '<span id="sp_' . $fieldName . '[PRODUCTS][' . $i . ']' . '">' . $productsNames[$products[$i]] . '</span></td></tr>';
                $index++;
            }
        }
        $html .= '</table> ';

        $html .= '<input type="button" value="Добавить..." 
        onclick="jsUtils.OpenWindow(\'/bitrix/admin/iblock_element_search.php?lang=' . LANGUAGE_ID . '&amp;IBLOCK_ID=5&amp;n=' . $fieldName . '&amp;m=y&amp;iblockfix=y\', 900, 700);">';
        $html .= '</div>';
        $html .= '</div><br/>';

        $html .= '<script>' . "\r\n";
        $html .= "var MV_" . md5($fieldName) . " = " . $index . ";\r\n";
        $html .= "function InS" . md5($fieldName) . "(id, name){ \r\n";
        $html .= "	oTbl=document.getElementById('dv" . md5($fieldName) . "');\r\n";
        $html .= "	oRow=oTbl.insertRow(oTbl.rows.length-1); \r\n";
        $html .= "	oCell=oRow.insertCell(-1); \r\n";
        $html .= "	oCell.innerHTML=" .
            "'&nbsp; <input name=\"" . $fieldName . "[PRODUCTS]['+MV_" . md5($fieldName) . "+']\" value=\"'+id+'\" id=\"" . $fieldName . "[PRODUCTS]['+MV_" . md5($fieldName) . "+']\" size=\"5\" type=\"text\">'+\r\n" .
            "'&nbsp; <input type=\"button\" value=\"...\" '+\r\n" .
            "'onClick=\"jsUtils.OpenWindow(\'/bitrix/admin/iblock_element_search.php?lang=" . LANGUAGE_ID . "&amp;IBLOCK_ID=5&amp;n=" . $fieldName . "[PRODUCTS]['+MV_" . md5($fieldName) . "+']&amp;iblockfix=y" . "\', '+\r\n" .
            "' 900, 700);\"> &nbsp;<span id=\"sp_" . $fieldName . "[PRODUCTS]['+MV_" . md5($fieldName) . "+']\" >'+name+'</span>" .
            "';";
        $html .= 'MV_' . md5($fieldName) . '++;';
        $html .= '}';
        $html .= "\r\n</script>";

        return $html;
    }
}