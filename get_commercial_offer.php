<?php require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");

use Bitrix\Iblock\ElementTable;
use Bitrix\Main\Loader;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Html;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Font;

$phpwordOn = include_once($_SERVER['DOCUMENT_ROOT'] . '/local/lib/vendor/autoload.php');

if ($phpwordOn && Loader::includeModule('iblock') && count($_GET) == 2) {

    $id = $_GET['id'];
    $blockId = $_GET['blockId'];

    $dbItems = ElementTable::getList(array(
        'filter' => array('IBLOCK_ID' => $blockId, 'ID' => $id)
    ));
    if ($arItem = $dbItems->fetch()) {
        $itemName = $arItem['NAME'];
        $filename = $itemName . ".docx";
        CModule::IncludeModule("catalog");
        $arPrice = CPrice::GetBasePrice($id);
        $price = number_format($arPrice['PRICE'], 0, ' ', ' ');
        $detailPhotoId = $arItem['DETAIL_PICTURE'];

        $detailPhotoSrc = $_SERVER['DOCUMENT_ROOT'] . CFile::GetPath($detailPhotoId);

        $result = CIBlockElement::getProperty(
            $arItem['IBLOCK_ID'],
            $arItem['ID'],
            [],
            [
                'CODE' => 'MORE_PHOTO',
            ]
        );

        while ($morePhotoElement = $result->Fetch()) {
            if (isset($morePhotoElement['VALUE']))
                $arMorePhoto[] = $morePhotoElement;
        }

        $result = CIBlockElement::getProperty(
            $arItem['IBLOCK_ID'],
            $arItem['ID'],
            [],
            [
                'CODE' => 'AVAILABILITY',
            ]
        );
        $availability = $result->Fetch();
        $availabilityValue = $availability['VALUE_ENUM'];

        $result = CIBlockElement::getProperty(
            $arItem['IBLOCK_ID'],
            $arItem['ID'],
            [],
            [
                'CODE' => 'TABLE',
            ]
        );
        $table = $result->Fetch();
        $tableText = $table['VALUE']['TEXT'];

        $previewText = strip_tags($arItem['PREVIEW_TEXT']);
        $detailText = $arItem['DETAIL_TEXT'];
        $defaultParagraphStyle = ['spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pointToTwip(0), 'spaceBefore' => \PhpOffice\PhpWord\Shared\Converter::pointToTwip(0), 'spacing' => 0];

        $phpWord = new PhpWord();
        $phpWord->setDefaultFontName('Arial');
        $phpWord->setDefaultFontSize(10);
        $phpWord->setDefaultParagraphStyle($defaultParagraphStyle);

        $section = $phpWord->addSection([
            'marginLeft' => 567,
            'marginRight' => 567,
            'marginTop' => 567,
            'marginBottom' => 567,
        ]);

        /*STYLES*/
        $headerStyle = ['size' => 9];
        $headerStyleLink = ['color' => '0000FF', 'underline' => Font::UNDERLINE_SINGLE, 'size' => 9];
        $font10 = ['size' => 10];

        /*HEADER*/
        $header = $section->addHeader();

        $tableBorderStyle = ['borderSize' => 'none'];

        $headerTable = $header->addTable($tableBorderStyle);
        $headerTable->addRow(300, $tableBorderStyle);

        $headerText = $headerTable->addCell(6500, $tableBorderStyle);

        $headerText->addText('ООО «САЙН СЕРВИС»', $headerStyle);
        $headerText->addText('г. Новосибирск, ул. Семьи Шамшиных, 99, оф. 409', $headerStyle);
        $headerText->addText('г. Зеленоград, проспект Генерала Алексеева, 42с1, 2 этаж, офис 233', $headerStyle);
        $headerText->addText('ИНН 5406802893 КПП 540601001 ОГРН 1195476086574', $headerStyle);

        $headerTextContacts = $headerText->addTextRun();
        $headerTextContacts->addText('Тел: 8 800 555 9419 E-mail:', $headerStyle);
        $headerTextContacts->addLink('mailto:info@sign-service.ru', 'info@sign-service.ru', $headerStyleLink);
        $headerTextContacts->addText(' Web: ', $headerStyle);
        $headerTextContacts->addLink('https://sign-service.ru', 'www.sign-service.ru', $headerStyleLink);

        $headerTable->addCell(4200, $tableBorderStyle)->addImage($_SERVER['DOCUMENT_ROOT'] . '/include/commercial_offer_files/logo.jpg', array(
            'width' => 388 / 2,
            'height' => 103 / 2,
            'align' => 'right'
        ));

        /*TITLE*/
        $section->addText($itemName, ['size' => 16, 'bold' => true], ['align' => 'center']);

        /*DETAIL PHOTO*/
        if (is_file($detailPhotoSrc)) {
            list($widthPhoto, $heightPhoto, $type, $attr) = getimagesize($detailPhotoSrc);
            $widthPhotoEnd = 300;
            $heightPhotoEnd = (300 / $widthPhoto) * $heightPhoto;
            $section->addImage($detailPhotoSrc, ['height' => $heightPhotoEnd, 'width' => $widthPhotoEnd, 'alignment' => 'center']);
        }

        /*AVAILABILITY*/
        $section->addText($availabilityValue, ['size' => 16, 'bold' => true], ['align' => 'left']);

        /*TABLE FROM PROPERTY*/
        $tableText = preg_replace('/(<[^>]+) style=".*?"/i', '$1', $tableText);
        $tableText = preg_replace("/\t\n|\t|\n/", '', $tableText);
        $tableText = trim(strip_tags($tableText, '<table><tr><td><b><strong><p>'));
        $tableText = str_replace("&nbsp;", '', $tableText);
        $tableText = preg_replace("/<tr[^>]*>(\s{0,}(<[a-z]*[^>]*>){0,}\s{0,}){0,}<\\/tr>/", '', $tableText);
        $tableText = preg_replace('|<tr>(.*)(</tr>)|Uis', '<tr style="font-weight: bold; font-size: 16pt ">$1</tr>', $tableText, 1);
        $tableText = preg_replace('/.<\/table>/',
            '$1<tr>
                            <td>
                                <p style="font-weight: bold; font-size: 16pt;text-align: center;">Цена, руб</p>
                            </td>
                            <td colspan="10">
                                <p style="font-weight: bold; font-size: 16pt;text-align: center;">' . $price . ' </p>
                            </td>
                         </tr></table>', $tableText);
        Html::addHtml($section, $tableText, false, false);

        /*FIX TEXT*/
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('Цена включает НДС и таможенные сборы, инсталляцию оборудования (включает: сборку, аппаратную настройку, настройку ПО, тестовая работа, обучение оператора).', $font10, ['align' => 'both']);
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('В цену не входит доставка оборудования со склада компании, командировочные инженера на инсталляцию.', $font10, ['align' => 'both']);
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('Гарантийный срок на оборудование - 12 месяцев, при инсталляции оборудования представителем нашей компании.', $font10, ['align' => 'both']);
        $section->addText('', $font10, ['align' => 'both']);

        /*DETAIL TEXT*/
        $detailText = preg_replace("/\t\n|\t|\n/", '', $detailText);
        $detailText = str_replace('<br>', "&nbsp;", $detailText);
        $detailText = strstr($detailText, "<h2>");
        Html::addHtml($section, $detailText, false, false);

        /*TABLE FROM PHOTO*/
        if (isset($arMorePhoto) and !empty($arMorePhoto)) {

            $section->addText('');

            $tableFromPhoto = $section->addTable(['borderSize' => 4, 'cellMargin' => 80]);

            $colNumber = 1;
            foreach ($arMorePhoto as $itemPhoto) {

                if ($colNumber == 1) {
                    $tableFromPhoto->addRow();
                }

                $cell = $tableFromPhoto->addCell(3700);

                $photoSrc = $_SERVER['DOCUMENT_ROOT'] . CFile::GetPath($itemPhoto['VALUE']);

                if (is_file($photoSrc)) {
                    list($widthPhoto, $heightPhoto, $type, $attr) = getimagesize($photoSrc);

                    $cell->addImage($photoSrc, array(
                        'width' => 165,
                        'height' => (165 / $widthPhoto) * $heightPhoto,
                        'align' => 'center'
                    ));
                }


                $cell->addText($itemPhoto['DESCRIPTION'], ['size' => 10], ['align' => 'center']);

                $colNumber = $colNumber == 3 ? 1 : $colNumber + 1;

            }
        }

        /*PREVIEW TEXT*/
        $linkStyle = ['color' => '0000FF', 'underline' => Font::UNDERLINE_SINGLE, 'size' => 10];

        $section->addText('');
		//Html::addHtml($section, $previewText, false, false);
        $section->addText($previewText, $font10, ['align'=>'both']);

        /*CONTACTS*/
        $section->addText('');

        $contactsPagStyle = ['align' => 'right', 'spaceBefore' => 0, 'spaceAfter' => 0];
        $section->addText('С уважением,', $font10, $contactsPagStyle);
        $section->addText('Компания «САЙН СЕРВИС»', $font10, $contactsPagStyle);
        $section->addText('Москва, г. Зеленоград, пр-т Генерала Алексеева, 42 ст1, 2 этаж, оф. 233', $font10, $contactsPagStyle);
        $section->addText('г. Новосибирск, ул. Семьи Шамшиных, 99, 4 этаж, офис 409', $font10, $contactsPagStyle);
        $section->addText('тел. +7 (800) 555-94-19', $font10, $contactsPagStyle);

        $contactsMail = $section->addTextRun($contactsPagStyle);
        $contactsMail->addText('e-mail ', $font10, $contactsPagStyle);
        $contactsMail->addLink('mailto:info@sign-service.ru', 'info@sign-service.ru', $linkStyle, $contactsPagStyle);
        $section->addLink('https://sign-service.ru', 'www.sign-service.ru', $linkStyle, $contactsPagStyle);

        /*SAVE TO TMP*/
        $objWriter = IOFactory::createWriter($phpWord);
        $temp_file_uri = tempnam('', 'xyz');

        $objWriter->save($temp_file_uri);

        /*GIVE TO USER*/
        header('Content-Description: File Transfer');
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename=' . $filename);
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Content-Length: ' . filesize($temp_file_uri));
        readfile($temp_file_uri);
        unlink($temp_file_uri); // deletes the temporary file
        exit;
    }
} ?>