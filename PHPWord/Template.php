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

    private $_headerXMLArr;
    private $_footerXMLArr;
    private $_rels;
    private $_types;
    private $_countRels;
    private $_customXML;
    private $_hFDirArr;

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
        $this->_rels        = $this->_objZip->getFromName('word/_rels/document.xml.rels');
        $this->_types       = $this->_objZip->getFromName('[Content_Types].xml');
        $this->_customXML   = $this->_objZip->getFromName('docProps/custom.xml');
        $this->_countRels   = substr_count($this->_rels, 'Relationship') - 1;

        //页眉文件读取
        $this->getPageHeaderAndFooterXml($this->_tempFileName); //页眉和页脚存在多个XML
        if (isset($this->_hFDirArr['header'])) {
            foreach ($this->_hFDirArr['header'] as $key => $headerDir) {
                $this->_headerXMLArr[$key]  = $this->_objZip->getFromName($headerDir);
            }
        } else {
            $this->_headerXMLArr[0]  = $this->_objZip->getFromName('word/header1.xml');
        }

        //页脚文件读取
        if (isset($this->_hFDirArr['footer'])) {
            foreach ($this->_hFDirArr['footer'] as $key => $footerDir) {
                $this->_footerXMLArr[$key]  = $this->_objZip->getFromName($footerDir);
            }
        } else {
            $this->_footerXMLArr[0]  = $this->_objZip->getFromName('word/footer1.xml');
        }
    }

    /**
     * 标签替换为文本 - 文档主体
     * @param string $search
     * @param string $replace
     */
    public function setValue($search, $replace) {
        $search = $this->lable($search);

        if(mb_detect_encoding($replace, mb_detect_order(), true) !== 'UTF-8') {
            $replace = utf8_encode($replace);
        }

        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);
    }

    /**
     * 标签替换为文本 - 页眉页脚
     * @param string $search
     * @param string $replace
     * @param string $type - 类别: header OR footer
     */
    public function setHeaderAndFooterValue($search, $replace, $type) {
        $search = $this->lable($search);

        if(mb_detect_encoding($replace, mb_detect_order(), true) !== 'UTF-8') {
            $replace = utf8_encode($replace);
        }

        if ($type == 'header') {
            foreach ($this->_headerXMLArr as $key => $item) {
                $this->_headerXMLArr[$key] = str_replace($search, $replace, $item);
            }
        }

        if ($type == 'footer') {
            foreach ($this->_footerXMLArr as $key => $item) {
                $this->_footerXMLArr[$key] = str_replace($search, $replace, $item);
            }
        }
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
            if (isset($img['imgExt'])) {
                $imgExt = $img['imgExt'];
            } else {
                $imgExt = substr($img['src'], strrpos($img['src'], '.')+1);
            }
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
        if ($rowStart === false || $rowEnd === 7) {
            return false;
        }

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
            $result .= preg_replace('/\${([^#]*?)\}/', '\${\1#' . $i . '}', $xmlRow);
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
        $this->_objZip->addFromString('word/_rels/document.xml.rels', $this->_rels);
        $this->_objZip->addFromString('[Content_Types].xml', $this->_types);
        $this->_objZip->addFromString('docProps/custom.xml', $this->_customXML);

        //页眉文件读取
        if (isset($this->_hFDirArr['header'])) {
            foreach ($this->_hFDirArr['header'] as $key => $headerDir) {
                $this->_objZip->addFromString($headerDir, $this->_headerXMLArr[$key]);
            }
        } else {
            $this->_objZip->addFromString('word/header1.xml', $this->_headerXMLArr[0]);
        }

        //页脚文件读取
        if (isset($this->_hFDirArr['footer'])) {
            foreach ($this->_hFDirArr['footer'] as $key => $footerDir) {
                $this->_objZip->addFromString($footerDir, $this->_footerXMLArr[$key]);
            }
        } else {
            $this->_objZip->addFromString('word/footer1.xml', $this->_footerXMLArr[0]);
        }

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
            //其他匹配规则
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

    /**
     *  获取页眉 - 页脚
     * @param string $fileDir - 文件路径
     * @return array
     */
    public function getPageHeaderAndFooterXml($fileDir)
    {
        $this->_hFDirArr = []; //页眉和页脚文件路径
        $zip = zip_open($fileDir);
        if ($zip) {
            while ($zip_entry = zip_read($zip)) {
                if (stripos(zip_entry_name($zip_entry), 'word/header') !== false) {
                    $this->_hFDirArr['header'][] = zip_entry_name($zip_entry);
                }

                if (stripos(zip_entry_name($zip_entry), 'word/footer') !== false) {
                    $this->_hFDirArr['footer'][] = zip_entry_name($zip_entry);
                }
            }
            zip_close($zip);
        }
    }

    /**
     * 获取 customXML - 文件属性
     */
    public function getCustomXML() {
        return $this->_customXML;
    }

    /**
     * 设置 customXML - 文件属性
     * @param string $customXML
     */
    public function setCustomXML($customXML) {
        $this->_customXML = $customXML;
    }
}
