<?php

namespace Maatwebsite\Excel;

use ArrayAccess;
use Closure;
use Illuminate\Support\Collection;
use PhpOffice\PhpSpreadsheet\Worksheet\Row as SpreadsheetRow;

/** @mixin SpreadsheetRow */
class Row implements ArrayAccess
{
    use DelegatedMacroable;

    /**
     * @var array
     */
    protected $headingRow = [];

    /**
     * @var \Closure
     */
    protected $preparationCallback;

    /**
     * @var SpreadsheetRow
     */
    protected $row;

    /**
     * @var array|null
     */
    protected $rowCache;

    /**
     * @var string
     */
    protected $rowCacheKey;

    /**
     * @param  SpreadsheetRow  $row
     * @param  array  $headingRow
     * @param  array  $headerIsGrouped
     */
    public function __construct(SpreadsheetRow $row, array $headingRow = [], array $headerIsGrouped = [])
    {
        $this->row             = $row;
        $this->headingRow      = $headingRow;
        $this->headerIsGrouped = $headerIsGrouped;
    }

    /**
     * @return SpreadsheetRow
     */
    public function getDelegate(): SpreadsheetRow
    {
        return $this->row;
    }

    public function getHeadingRow(): array
    {
        return $this->headingRow;
    }

    /**
     * @param string $startColumn
     * @param string|null $endColumn
     * @return array<Cell|Cell[]>
     */
    public function getCells(string $startColumn = 'A', ?string $endColumn = null): array
    {
        $cells = [];

        $i = 0;
        foreach ($this->row->getCellIterator($startColumn, $endColumn) as $cell) {
            $cell = new Cell($cell);

            if (isset($this->headingRow[$i])) {
                if ($this->headerIsGrouped[$i]) {
                    $cells[$this->headingRow[$i]][] = $cell;
                } else {
                    $cells[$this->headingRow[$i]] = $cell;
                }
            } else {
                $cells[] = $cell;
            }

            $i++;
        }

        return $cells;
    }

    /**
     * @param mixed $nullValue
     * @param bool $calculateFormulas
     * @param bool $formatData
     * @param string|null $endColumn
     * @return Collection
     */
    public function toCollection($nullValue = null, bool $calculateFormulas = false, bool $formatData = true, ?string $endColumn = null): Collection
    {
        return new Collection($this->toArray($nullValue, $calculateFormulas, $formatData, $endColumn));
    }

    /**
     * @param mixed $nullValue
     * @param bool $calculateFormulas
     * @param bool $formatData
     * @param string|null $endColumn
     * @return array
     */
    public function toArray($nullValue = null, bool $calculateFormulas = false, bool $formatData = true, ?string $endColumn = null): ?array
    {
        $serializedArguments = serialize(func_get_args());

        if ($serializedArguments === $this->rowCacheKey && is_array($this->rowCache)) {
            return $this->rowCache;
        }

        $values = $this->getCellsValues($this->getCells('A', $endColumn), $nullValue, $calculateFormulas, $formatData);

        if (isset($this->preparationCallback)) {
            $values = ($this->preparationCallback)($values, $this->row->getRowIndex());
        }

        $this->rowCacheKey = $serializedArguments;
        $this->rowCache = $values;

        return $values;
    }

    /**
     * @param array<Cell|Cell[]> $cells
     * @param mixed $nullValue
     * @param bool $calculateFormulas
     * @param bool $formatData
     * @return array
     */
    protected function getCellsValues(array $cells, $nullValue = null, bool $calculateFormulas = false, bool $formatData = true): array
    {
        return array_map(function($cell) use ($nullValue, $calculateFormulas, $formatData) {
            if (is_array($cell)) {
                return $this->getCellsValues($cell, $nullValue, $calculateFormulas, $formatData);
            }
            return $cell->getValue($nullValue, $calculateFormulas, $formatData);
        }, $cells);
    }

    /**
     * @param bool $calculateFormulas
     * @param  string|null  $endColumn
     * @return bool
     */
    public function isEmpty(bool $calculateFormulas = false, bool $formatData = true, ?string $endColumn = null): bool
    {
        return count(array_filter($this->toArray(null, $calculateFormulas, $formatData, $endColumn))) === 0;
    }

    /**
     * @return int
     */
    public function getIndex(): int
    {
        return $this->row->getRowIndex();
    }

    /**
     * @param Closure|null $preparationCallback
     *
     * @internal
     */
    public function setPreparationCallback(Closure $preparationCallback = null): void
    {
        $this->preparationCallback = $preparationCallback;
    }

    protected function getCached(): ?array
    {
        return $this->rowCache ?: $this->toArray();
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        return isset(($this->getCached())[$offset]);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        return ($this->getCached())[$offset];
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        //
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        //
    }
}
