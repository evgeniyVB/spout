<?php

namespace Box\Spout\Common\Entity\Style;

/**
 * Class Alignment
 * This class provides constants to work with text alignment.
 */
abstract class CellVerticalAlignment
{
    const TOP = 'top';
    const BOTTOM = 'bottom';
    const MIDDLE = 'center';

    private static $VALID_ALIGNMENTS = [
        self::TOP => 1,
        self::BOTTOM => 1,
        self::MIDDLE => 1,
    ];

    /**
     * @param string $cellAlignment
     *
     * @return bool Whether the given cell alignment is valid
     */
    public static function isValid($cellAlignment)
    {
        return isset(self::$VALID_ALIGNMENTS[$cellAlignment]);
    }
}
