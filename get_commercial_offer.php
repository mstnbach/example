<?php require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");

use Bitrix\Iblock\ElementTable;
use Bitrix\Main\Application;
use Bitrix\Main\Loader;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Html;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Font;

$phpwordOn = include_once($_SERVER['DOCUMENT_ROOT'] . '/local/lib/vendor/autoload.php');
if ($phpwordOn && Loader::includeModule('iblock')) {

    $id = $_POST['id'];
    $blockId = $_POST['blockId'];
    $confData = $_POST['confData'];

    $dbItems = ElementTable::getList(array(
        'filter' => array('IBLOCK_ID' => $blockId, 'ID' => $id)
    ));
    if ($arItem = $dbItems->fetch()) {
        $itemName = $arItem['NAME'];
        $fileName = $itemName . ".docx";

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

        $confResult = [];

        if (!empty($confData)) {

            $confAr = json_decode($confData);

            $confMaxData = [];
            $confMaxResult = CIBlockElement::getProperty(
                $arItem['IBLOCK_ID'],
                $arItem['ID'],
                [],
                [
                    'CODE' => 'CONFIGURATION',
                ]
            );

            while ($confMaxResultData = $confMaxResult->Fetch()) {
                $confValue = unserialize($confMaxResultData['VALUE']);
                $confMaxData[$confValue['GROUP']] = $confValue['MAX_COUNT'];
            }

            foreach ($confAr as $confItem) {

                $confItem = get_object_vars($confItem);

                $maxCount = empty($confMaxData[$confItem['GROUP_ID']]) ? 0 : $confMaxData[$confItem['GROUP_ID']];

                $quantity = $confItem['QUANTITY'] > $maxCount && $maxCount > 0 ? $maxCount : $confItem['QUANTITY'];

                $confResult[$confItem['SET_PRODUCT']] = [
                    'QUANTITY' => $quantity
                ];
            }
            $confProdIds = array_keys($confResult);

            $idTypePrice = CCatalogGroup::GetBaseGroup();

            $prodProps = CIBlockElement::GetList(
                [],
                ['ID' => $confProdIds],
                false,
                false,
                ['ID', 'NAME', 'CATALOG_PRICE_' . $idTypePrice['ID']]
            );

            while ($prod = $prodProps->Fetch()) {
                $confResult[$prod['ID']]['NAME'] = $prod['NAME'];
                $confResult[$prod['ID']]['PRICE'] = $prod['CATALOG_PRICE_' . $idTypePrice['ID']];
            }
        }

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

        $headerText->addText('?????? ?????????? ??????????????', $headerStyle);
        $headerText->addText('??. ??????????????????????, ????. ?????????? ????????????????, 99, ????. 409', $headerStyle);
        $headerText->addText('??. ????????????????????, ???????????????? ???????????????? ??????????????????, 42??1, 2 ????????, ???????? 233', $headerStyle);
        $headerText->addText('?????? 5406802893 ?????? 540601001 ???????? 1195476086574', $headerStyle);

        $headerTextContacts = $headerText->addTextRun();
        $headerTextContacts->addText('??????: 8 800 555 9419 E-mail:', $headerStyle);
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
                                <p style="font-weight: bold; font-size: 16pt;text-align: center;">????????, ??????</p>
                            </td>
                            <td colspan="10">
                                <p style="font-weight: bold; font-size: 16pt;text-align: center;">' . $price . ' </p>
                            </td>
                         </tr></table>', $tableText);
        Html::addHtml($section, $tableText, false, false);

        /*FIX TEXT*/
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('???????? ???????????????? ?????? ?? ???????????????????? ??????????, ?????????????????????? ???????????????????????? (????????????????: ????????????, ???????????????????? ??????????????????, ?????????????????? ????, ???????????????? ????????????, ???????????????? ??????????????????).', $font10, ['align' => 'both']);
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('?? ???????? ???? ???????????? ???????????????? ???????????????????????? ???? ???????????? ????????????????, ?????????????????????????????? ???????????????? ???? ??????????????????????.', $font10, ['align' => 'both']);
        $section->addText('', $font10, ['align' => 'both']);
        $section->addText('?????????????????????? ???????? ???? ???????????????????????? - 12 ??????????????, ?????? ?????????????????????? ???????????????????????? ???????????????????????????? ?????????? ????????????????.', $font10, ['align' => 'both']);
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
        $section->addText($previewText, $font10, ['align' => 'both']);

        /*TABLE OF CONFIGURATION*/

        if (!empty($confResult)) {

            $section->addText('');
            $section->addText('?????????????? ???? ??????????????????????????', ['size' => 16, 'bold' => true], ['align' => 'center']);
            $section->addText('');

            $totalSum = $arPrice['PRICE'];

            $confHtml = '<table cellspacing="0" cellpadding="2" border="1" align="center" ><tr><th width="60%">??????????</th><th width="20%">????????????????????</th><th width="20%">???????? ???? ??????????????, ??????</th></tr>';

            foreach ($confResult as $confItem) {
                $totalSum += $confItem['PRICE'] * $confItem['QUANTITY'];
                $confHtml .= '<tr align="center"><td>' . $confItem['NAME'] . '</td><td>' . $confItem['QUANTITY'] . '</td><td>' . number_format($confItem['PRICE'], 0, ' ', ' ') . '</td></tr>';
            }
            $confHtml .= '</table>';

            Html::addHtml($section, $confHtml, false, false);

            $section->addText('');
            $section->addText('?????????? ??????????????????: '
                . number_format($totalSum, 0, ' ', ' ')
                . ' ??????.', ['size' => 16, 'bold' => true], ['align' => 'center']);
            $section->addText('');
        }

        /*CONTACTS*/
        $section->addText('');

        $contactsPagStyle = ['align' => 'right', 'spaceBefore' => 0, 'spaceAfter' => 0];
        $section->addText('?? ??????????????????,', $font10, $contactsPagStyle);
        $section->addText('???????????????? ?????????? ??????????????', $font10, $contactsPagStyle);
        $section->addText('????????????, ??. ????????????????????, ????-?? ???????????????? ??????????????????, 42 ????1, 2 ????????, ????. 233', $font10, $contactsPagStyle);
        $section->addText('??. ??????????????????????, ????. ?????????? ????????????????, 99, 4 ????????, ???????? 409', $font10, $contactsPagStyle);
        $section->addText('??????. +7 (800) 555-94-19', $font10, $contactsPagStyle);

        $contactsMail = $section->addTextRun($contactsPagStyle);
        $contactsMail->addText('e-mail ', $font10, $contactsPagStyle);
        $contactsMail->addLink('mailto:info@sign-service.ru', 'info@sign-service.ru', $linkStyle, $contactsPagStyle);
        $section->addLink('https://sign-service.ru', 'www.sign-service.ru', $linkStyle, $contactsPagStyle);

        /*SAVE TO TMP*/
        $objWriter = IOFactory::createWriter($phpWord);
        $root = Application::getDocumentRoot();

        CheckDirPath($_SERVER["DOCUMENT_ROOT"] . '/bitrix/tmp/offers');

        $tempFileUri = tempnam($root . '/offers', 'offer');

        $objWriter->save($tempFileUri);

        $fileUri = strpos($tempFileUri, 'offers') ? str_replace($root, '', $tempFileUri) : $tempFileUri;

        /*GIVE TO USER*/
        $fileResult['fileUri'] = $fileUri;
        $fileResult['fileName'] = $fileName;

        exit(json_encode($fileResult));
    }
} ?>