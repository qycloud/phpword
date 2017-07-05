<?php
/**
 * 此为基于官网【0.6.2-1 Beta】改良版
 * 改动：
 *      １．修正UTF8乱码问题
 *      ２．修正标签残缺导致替换失败
 *      ３．支持标签替换为克隆表格
 *      ４．支持标签替换时导入图片
 * http://phpword.codeplex.com/PHPWord
 * http://my.oschina.net/cart
 * @author LET(QQ:498936940)
 *
 */
class PHPWord_Template {
    private $_objZip;
    private $_tempFileName;
    private $_documentXML;

    private $_header1XML;
    private $_footer1XML;
    private $_rels;
    private $_types;
    private $_countRels;

    /**
     * 初始化Word模板
     * @param string $strFilename
     */
    public function __construct($strFilename) {
        $path = dirname($strFilename);
        $this->_tempFileName = $path.DIRECTORY_SEPARATOR.time().'.docx';

        copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

        $this->_objZip = new ZipArchive();
        $this->_objZip->open($this->_tempFileName);

        $this->_documentXML = $this->fixBrokenMacros($this->_objZip->getFromName('word/document.xml'));

        $this->_header1XML  = $this->_objZip->getFromName('word/header1.xml');
        $this->_footer1XML  = $this->_objZip->getFromName('word/footer1.xml');
        $this->_rels        = $this->_objZip->getFromName('word/_rels/document.xml.rels');
        $this->_types       = $this->_objZip->getFromName('[Content_Types].xml');
        $this->_countRels   = substr_count($this->_rels, 'Relationship') - 1;
    }

    /**
     * 标签替换为文本
     * @param string $search
     * @param string $replace
     */
    public function setValue($search, $replace) {
        $search = $this->lable($search);

        if(mb_detect_encoding($replace, mb_detect_order(), true) !== 'UTF-8') {
            $replace = utf8_encode($replace);
        }

        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);

        $this->_header1XML = str_replace($search, $replace, $this->_header1XML);
        $this->_footer1XML = str_replace($search, $replace, $this->_footer1XML);
    }

    /**
     * 标签替换为图片
     * @param string $search
     * @param array $replace
     */
    public function setImage($search, array $replace){
        $relationTmpl = '<Relationship Id="RID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/IMG"/>';
        $imgTmpl = '<w:pict><v:shape type="#_x0000_t75" style="width:WIDpx;height:HEIpx"><v:imagedata r:id="RID" o:title=""/></v:shape></w:pict>';
        $typeTmpl = ' <Override PartName="/word/media/IMG" ContentType="image/EXT"/>';
        $toAdd = $toAddImg = $toAddType = '';

        foreach($replace as $img){
            $imgExt = substr($img['src'], strrpos($img['src'], '.')+1);
            if (strtolower($imgExt) === 'jpg') {
                $imgExt = 'jpeg';
            }
            $imgName = 'img'.$this->_countRels.'.'.$imgExt;
            $rid = 'rId'.$this->_countRels++;

            $this->_objZip->addFile($img['src'], 'word/media/'.$imgName);

            $width = 400;
            $height = 200;
            if (isset($img['size'][1])){
                list($width, $height) = $img['size'];
            }

            $toAddImg .= str_replace(['RID', 'WID', 'HEI'], [$rid, $width, $height], $imgTmpl);
            if (isset($img['alt'])) {
                $toAddImg .= '<w:br/><w:t>' . $this->formatToXml($img['alt']) . '</w:t><w:br/>';
            }
            $toAddType .= str_replace(['IMG', 'EXT'], [$imgName, $imgExt], $typeTmpl) ;
            $toAdd .= str_replace(['RID', 'IMG'], [$rid, $imgName], $relationTmpl);
        }

        $this->_documentXML = str_replace('<w:t>${'.$search.'}</w:t>', $toAddImg, $this->_documentXML);
        $this->_types       = str_replace('</Types>', $toAddType, $this->_types) . '</Types>';
        $this->_rels        = str_replace('</Relationships>', $toAdd, $this->_rels) . '</Relationships>';
    }

    /**
     * 更新Word中的图片
     * @param string $path
     * @param string $imageName
     */
    public function replaceImage($path, $imageName) {
        $this->_objZip->deleteName('word/media/'.$imageName);
        $this->_objZip->addFile($path, 'word/media/'.$imageName);
    }

    /**
     * 根据标签克隆当前标签所在的整个表格行
     * @param string $search
     * @param int $numberOfClones
     * @throws Exception
     */
    public function cloneRow($search, $numberOfClones)
    {
        $search = $this->lable($search);

        $tagPos = strpos($this->_documentXML, $search);
        if (!$tagPos) {
            return false;
        }

        $rowStart = $this->findRowStart($tagPos);
        $rowEnd = $this->findRowEnd($tagPos);
        $xmlRow = $this->getSlice($rowStart, $rowEnd);

        if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow)) {
            // $extraRowStart = $rowEnd;
            $extraRowEnd = $rowEnd;
            while (true) {
                $extraRowStart = $this->findRowStart($extraRowEnd + 1);
                $extraRowEnd = $this->findRowEnd($extraRowEnd + 1);

                if ($extraRowEnd < 7) {
                    break;
                }

                $tmpXmlRow = $this->getSlice($extraRowStart, $extraRowEnd);
                if (!preg_match('#<w:vMerge/>#', $tmpXmlRow) &&
                    !preg_match('#<w:vMerge w:val="continue" />#', $tmpXmlRow)) {
                        break;
                    }
                    $rowEnd = $extraRowEnd;
            }
            $xmlRow = $this->getSlice($rowStart, $rowEnd);
        }

        $result = $this->getSlice(0, $rowStart);
        for ($i = 0; $i <= $numberOfClones; $i++) {
            $result .= preg_replace('/\$\{(.*?)\}/', '\${\\1#' . $i . '}', $xmlRow);
        }
        $result .= $this->getSlice($rowEnd);

        $this->_documentXML = $result;
    }

    /**
     * 保存模板
     * @param string $strFilename
     * @throws Exception
     */
    public function save($strFilename) {
        if(file_exists($strFilename)) {
            unlink($strFilename);
        }

        $this->_objZip->addFromString('word/document.xml', $this->_documentXML);

        $this->_objZip->addFromString('word/header1.xml', $this->_header1XML);
        $this->_objZip->addFromString('word/footer1.xml', $this->_footer1XML);
        $this->_objZip->addFromString('word/_rels/document.xml.rels', $this->_rels);
        $this->_objZip->addFromString('[Content_Types].xml', $this->_types);

        // Close zip file
        if($this->_objZip->close() === false) {
            throw new Exception('Could not close zip file.');
        }

        rename($this->_tempFileName, $strFilename);
    }

    private function findRowStart($offset)
    {
        $rowStart = strrpos($this->_documentXML, '<w:tr ', ((strlen($this->_documentXML) - $offset) * -1));

        if (!$rowStart) {
            $rowStart = strrpos($this->_documentXML, '<w:tr>', ((strlen($this->_documentXML) - $offset) * -1));
        }
        if (!$rowStart) {
            //throw new Exception('Can not find the start position of the row to clone.');
        }

        return $rowStart;
    }

    private function findRowEnd($offset)
    {
        return strpos($this->_documentXML, '</w:tr>', $offset) + 7;
    }

    private function getSlice($startPosition, $endPosition = 0)
    {
        if (!$endPosition) {
            $endPosition = strlen($this->_documentXML);
        }

        return substr($this->_documentXML, $startPosition, ($endPosition - $startPosition));
    }

    private function fixBrokenMacros($documentPart)
    {
        $fixedDocumentPart = $documentPart;

        $fixedDocumentPart = preg_replace_callback(
            '|\$[^{]*\{[^}]*\}|U',
            function ($match) {
                return strip_tags($match[0]);
            },
            $fixedDocumentPart
        );

        return $fixedDocumentPart;
    }

    private function lable($search)
    {
        if ('${' !== substr($search, 0, 2) && '}' !== substr($search, -1)) {
            $search = '${'.$search.'}';
        }

        return $search;
    }

    private function formatToXml($str) {
        return str_replace(['&', '<', '>', "\n"], ['&amp;', '&lt;', '&gt;', "\n".'<w:br/>'], $str);
    }
}
