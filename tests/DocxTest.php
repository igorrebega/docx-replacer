<?php

namespace IRebega\DocxReplacer\Tests;

use IRebega\DocxReplacer\Docx;
use PHPUnit\Framework\TestCase;

class DocxTest extends TestCase
{
    /** @test */
    public function it_replace_text()
    {
        if (file_exists(__DIR__ . "/tmp/process.docx")) {
            unlink(__DIR__ . "/tmp/process.docx");
        }

        copy(__DIR__ . "/files/case1/file.docx", __DIR__ . "/tmp/process.docx");

        $docx = new Docx(__DIR__ . "/tmp/process.docx");
        $docx->replaceText('{{replace this}}', 'replaced');

        $this->assertEquals(
            filesize(__DIR__ . "/tmp/process.docx"),
            filesize(__DIR__ . "/files/case1/result.docx")
        );
    }

    /** @test */
    public function it_replace_right_text()
    {
        if (file_exists(__DIR__ . "/tmp/process.docx")) {
            unlink(__DIR__ . "/tmp/process.docx");
        }

        copy(__DIR__ . "/files/case2/file.docx", __DIR__ . "/tmp/process.docx");

        $docx = new Docx(__DIR__ . "/tmp/process.docx");
        $docx->replaceText('{{replace this}}', 'replaced');

        $this->assertEquals(
            filesize(__DIR__ . "/tmp/process.docx"),
            filesize(__DIR__ . "/files/case2/result.docx")
        );
    }
}
