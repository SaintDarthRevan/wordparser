<?php

require 'vendor/autoload.php';

$docs = glob('docs/*.docx');

$i = 0;
$articlesListContent = '';

foreach($docs as $doc) {
    $file_name = explode('/', $doc)[1];
    $file_name = explode('.d', $file_name)[0];

    $content = getContentFromWordDoc($file_name, ++$i);
    createHtml((string)$i, $content);
    $articlesListContent .= '<a href="'.$i.'.html">'.$content['title'].'</a><br />';
}

createHtml('Список статей', ['title' => 'Список статей', 'body' => $articlesListContent]);

function getContentFromWordDoc($file_name)
{
    $data = [
        'title' => '',
        'body' => '',
    ];

    $Reader = \PhpOffice\PhpWord\IOFactory::createReader('Word2007');
    $content = $Reader->load('docs/'.$file_name.'.docx');

    foreach($content->getSections() as $section) {
        $elems = $section->getElements();
        $first = true;
        foreach($elems as $e) {
            if (get_class($e) === 'PhpOffice\PhpWord\Element\TextRun') {
                $eContent = '';
                foreach($e->getElements() as $text) {
                    $font = $text->getFontStyle();

                    if ( $font->getSize()) {
                        $size = $font->getSize()/10;
                    } else  {
                        $size = 1;
                    }
                    $bold = $font->isBold() ? 'font-weight:700;' :'';
                    $color = $font->getColor();
                    $fontFamily = $font->getName();

                    $eContent .= '<span style="font-size:' . $size . 'em;font-family:' . $fontFamily . '; '.$bold.'; color:#'.$color.'">';
                    $eContent .= nl2br($text->getText()).'</span>';
                }
                if ($first == true) {
                    $data['title'] = strip_tags($eContent);
                    $data['body'] .= '<h1>'.$data['title'].'</h1>';
                    $first = false;
                } else {
                    $data['body'] .= '<p>'.$eContent.'</p>';
                }
            } elseif(get_class($e) === 'PhpOffice\PhpWord\Element\TextBreak') {
                $data['body'] .= '<br />';
            } else {
                $data['body'] .= nl2br($e->getText());
            }
        }
    }

    return $data;
}

function createHtml($filename, $data)
{
    $content = '<html>
    <head>
    <title>'.$data['title'].'</title>
    <meta name="title" content="'.$data['title'].'" />
    <meta name="description" content="'.$data['title'].' в Москве. Бесплатная диагностика. Гарантия качества. Записаться - 8(495)150-70-69" />
    <meta name="keywords" content="'.$data['title'].'" />
    </head>
    <body>
        '.$data['body'].'
    </body>
    </html>';

    $file = fopen('html/'.$filename.'.html', 'w');
    fwrite($file, $content);
    fclose($file);
}