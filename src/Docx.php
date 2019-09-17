<?php

namespace IRebega\DocxReplacer;

use IRebega\DocxReplacer\Exceptions\DocxException;

/**
 * Replace text to text or even text to images in DOCX files
 *
 * @author i.rebega <igorrebega@gmail.com>
 */
class Docx extends \ZipArchive
{
    // Use to change PX to EMU
    const PX_TO_EMU = 8625;
    const REL_LOCATION = 'word/_rels/document.xml.rels';
    const DOCUMENT_BODY_LOCATION = 'word/document.xml';
    const HEADER_LOCATION = 'word/header1.xml';
    const FOOTER_LOCATION = 'word/footer1.xml';

    /**
     * @var string
     */
    protected $path;

    /**
     * Docx constructor.
     *
     * @param string $path path to DOCX file
     *
     * @throws \Exception
     */
    public function __construct($path)
    {
        $this->path = $path;

        if ($this->open($path, \ZipArchive::CREATE) !== true) {
            throw new DocxException("Unable to open <$path>");
        }
    }

    /**
     * Replace one text to another
     * Case sensitive
     *
     * @param $from
     * @param $to
     */
    public function replaceText($from, $to)
    {
        $this->replaceTextInLocation($from, $to, self::HEADER_LOCATION);
        $this->replaceTextInLocation($from, $to, self::FOOTER_LOCATION);
        $this->replaceTextInLocation($from, $to, self::DOCUMENT_BODY_LOCATION);
    }

    /**
     * Will replace text with case sensitivity
     *
     * @param $from
     * @param $to
     */
    public function replaceTextInsensitive($from, $to)
    {
        $this->replaceTextInLocation($from, $to, self::HEADER_LOCATION, true);
        $this->replaceTextInLocation($from, $to, self::FOOTER_LOCATION, true);
        $this->replaceTextInLocation($from, $to, self::DOCUMENT_BODY_LOCATION, true);
    }

    /**
     * Replace text to given image
     *
     * @param string $text What text search
     * @param string $path Image to which we want to replace text
     *
     * @throws \Exception
     */
    public function replaceTextToImage($text, $path)
    {
        if (! file_exists($path)) {
            throw new \Exception('Image not exists');
        }
        list($width, $height, $type) = getimagesize($path);
        $ext = pathinfo($path, PATHINFO_EXTENSION);
        $name = StringHelper::random(10) . '.' . $ext;
        $zipPath = 'word/media/' . $name;
        $this->addFromString($zipPath, file_get_contents($path));

        $relId = $this->addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            "media/$name");

        $block = $this->getImageBlock($relId, $width, $height);

        $this->replaceTextToBlock($text, $block);
    }

    /**
     * Replace one text to another in $location
     *
     * @param $from
     * @param $to
     * @param $location
     * @param $caseInsensitive
     */
    private function replaceTextInLocation($from, $to, $location, $caseInsensitive = false)
    {
        $to = $this->fixLineBreaksInTo($to);

        $message = $this->getFromName($location);
        if ($caseInsensitive) {
            $message = str_ireplace($from, $to, $message);
        } else {
            $message = str_replace($from, $to, $message);
        }

        $this->addFromString($location, $message);

        $this->save();
    }

    /**
     * Fix line breaks issue
     * more details -  https://github.com/PHPOffice/PHPWord/issues/553
     *
     * @param string $to
     *
     * @return string
     */
    private function fixLineBreaksInTo($to)
    {
        return preg_replace('~\R~u', '</w:t><w:br/><w:t>', $to);
    }

    /**
     * Save changes to archive
     */
    private function save()
    {
        $this->close();
        $this->open($this->path, \ZipArchive::CREATE);
    }

    /**
     * This block we use to insert into document xml
     *
     * @param $relId
     * @param $width
     * @param $height
     *
     * @return mixed
     */
    private function getImageBlock($relId, $width, $height)
    {
        $block = file_get_contents(__DIR__ . '/../templates/image.xml');
        $block = str_replace('{RID}', $relId, $block);
        $block = str_replace('{WIDTH}', $width * self::PX_TO_EMU, $block);

        return str_replace('{HEIGHT}', $height * self::PX_TO_EMU, $block);
    }

    /**
     * Find block that have $text inside and replace that block to $customBlock
     *
     * @param $text
     * @param $block
     */
    private function replaceTextToBlock($text, $block)
    {
        $file = $this->getFromName(self::DOCUMENT_BODY_LOCATION);

        $pattern = "/<w:r[ >](?:(?!<w:r>|<\/w:r>).)+" . preg_quote($text, '/') . "(?:(?!<w:r>|<\/w:r>).)+<\/w:r>/s";
        $file = preg_replace($pattern, $block, $file);

        $this->addFromString(self::DOCUMENT_BODY_LOCATION, $file);
        $this->save();
    }

    /**
     * @param $type
     * @param $target
     *
     * @return string RelId
     */
    private function addRel($type, $target)
    {
        $file = $this->getFromName(self::REL_LOCATION);
        $xml = new \SimpleXMLElement($file);

        $lastId = $this->getLastMaxRelId($xml);

        $child = $xml->addChild("Relationship");
        $child->addAttribute('Id', 'rId' . ($lastId + 1));
        $child->addAttribute('Type', $type);
        $child->addAttribute('Target', $target);

        $this->addFromString(self::REL_LOCATION, $xml->asXML());

        return 'rId' . ($lastId + 1);
    }

    /**
     * @param \SimpleXMLElement $xml
     *
     * @return int
     */
    private function getLastMaxRelId(\SimpleXMLElement $xml)
    {
        $rel = $xml->Relationship[0]['Id'];
        $max = $this->getNumberFromRelId($rel);
        foreach ($xml->Relationship as $relationship) {
            $number = $this->getNumberFromRelId($relationship['Id']);
            if ($number > $max) {
                $max = $number;
            }
        }

        return $max;
    }

    /**
     * @param $relId
     *
     * @return int
     */
    private function getNumberFromRelId($relId)
    {
        preg_match('!\d+!', $relId, $matches);

        return (int) $matches[0];
    }
}
