<?php

namespace IRebega\DocxReplacer;

/**
 * Find text block in DOCX xml and replace it by custom block
 *
 * @author i.rebega <igorrebega@gmail.com>
 */
class ReplaceTextBlockToCustom
{
    /**
     * @var string
     */
    private $document;

    /**
     * ReplaceTextBlockToCustom constructor.
     *
     * @param $document string XML file text
     */
    public function __construct($document)
    {
        $this->document = $document;
    }

    /**
     * Find block that have $text inside and replace that block to $customBlock
     *
     * @param string $text
     * @param string $customBlock
     * @return string text with result
     */
    public function result($text, $customBlock)
    {
        while ($textBlock = $this->findBlockWithText($text)) {
            $this->document = str_replace($textBlock, $customBlock, $this->document);
        }

        return $this->document;
    }

    /**
     * Find <w:r> </w:r> block with $text inside
     * If success will return that block, if not => false
     *
     * @param $text
     * @return bool|string
     */
    private function findBlockWithText($text)
    {
        $blockStart = 0;
        $blockEnd = 0;

        $haveWord = false;

        for ($i = 1; $i < strlen($this->document) - 14; $i++) {
            if ($this->stringHaveWord($this->document, $i, '<w:r ')) {
                $blockStart = $i;
            }

            if ($this->stringHaveWord($this->document, $i, $text)) {
                $haveWord = true;
            }

            if ($this->stringHaveWord($this->document, $i, '</w:r>') && $haveWord) {
                $blockEnd = $i + 6;
                break;
            }
        }

        if ($blockEnd == 0) {
            return false;
        }

        return substr($this->document, $blockStart, ($blockEnd - $blockStart));
    }

    /**
     * If $string started from $startFrom have $word
     *
     * @param $string string
     * @param $startFrom int
     * @param $word string
     * @return bool
     */
    private function stringHaveWord($string, $startFrom, $word)
    {
        for ($j = 0; $j < strlen($word); $j++) {
            if ($string[$startFrom + $j] != $word[$j]) {
                return false;
            }
        }
        return true;
    }
}