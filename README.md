# vba-polysort
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for polymorphic sorting and searching through the `ICompare` interface: sort any array, also an array of objects, with a pluggable comparison.

Uses a **Dual-Pivot QuickSort + InsertionSort** (threshold 47, after Java's `DualPivotQuicksort`). Built-in comparers handle numbers and dates, and strings with binary or text comparison; a class that implements `ICompare` handles anything else. For arrays of values only, [vba-qsort](https://github.com/vgrstn/vba-qsort) is simpler.

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `PolySort.bas` | Module | `Sort`, `Search`, `IsSorted` and `CompareCustom` — the entry point |
| `ICompare.cls` | Interface | `ECompare` and `Assign`, `Compare`, `Equal`, `Swap` |
| `CompareDefault.cls` | Class | Comparison by VB operator — numbers, dates |
| `CompareBinary.cls` | Class | Case-sensitive string comparison (`vbBinaryCompare`) |
| `CompareText.cls` | Class | Case-insensitive string comparison (`vbTextCompare`) |
| `Person.cls` | Class | Example class that implements `ICompare` (by last name, prefix, first name) |
| `Tests/TestPolySort.bas` | Tests | 24 Rubberduck unit tests, one case per test |

Each file has a `_WithAttributes` version with the VB attributes baked in and the Rubberduck annotations removed; import those if you don't use Rubberduck. Import `ICompare` and all its implementations together.

---

## ⚙️ Public Interface

### `PolySort` module

| Member | Description |
|---|---|
| `Sort arr [, idx [, method [, asc]]]` | Sorts `arr` in place, or by index if `idx` is passed. `asc = False` sorts descending. |
| `Search(arr, value [, idx [, method [, start]]])` | Binary search in a sorted array. Returns the lowest matching position, or `Null`. Pass a previous result as `start` to find the next equal value. |
| `IsSorted(arr [, idx [, method]])` | Returns `True` if `arr` is sorted (ascending or descending). |
| `CompareCustom` *(Get/Set)* | The `ICompare` to use; `Nothing` restores the automatic choice. |

The automatic choice of the comparer (when `CompareCustom` is `Nothing`):

| Array | Comparer |
|---|---|
| numbers and dates | `CompareDefault` |
| strings, `method` omitted or `vbBinaryCompare` | `CompareBinary` |
| strings, `vbTextCompare` | `CompareText` |
| strings, `vbDatabaseCompare` | `CompareDefault` (by operator: binary in `PolySort`) |
| objects, UDTs, mixed values | error 17 "Custom compare not implemented" |

`Sort` and `IsSorted` check every element of a Variant array; `Search` only looks at the type of the array or of its first element, so that it stays a binary search. In the same way `Search` only checks the length of an index array (an index out of range raises error 9), while `IsSorted` checks every index.

### `ICompare` interface

| Method | Description |
|---|---|
| `Assign(variable, value)` | Assigns a value or an object reference (`Let` or `Set`) |
| `Compare(value1, value2)` | Returns `ecLess` (-1), `ecEqual` (0) or `ecGreater` (1) |
| `Equal(value1, value2)` | Returns `True` if the two values are equal |
| `Swap(variable1, variable2)` | Swaps the contents of two variables |

All parameters are `ByRef`: a comparison is made at every step of a sort, and `ByVal` would copy both values every time. An implementation must not change `value`, `value1` or `value2`.

---

## 🚀 Quick Start

```vb
Dim a() As Variant
a = Array(5, 3, 8, 1, 4)
Sort a                            ' -> (1, 3, 4, 5, 8)  CompareDefault

Dim s() As Variant
s = Array("Banana", "apple", "Cherry")
Sort s, , vbTextCompare           ' -> ("apple", "Banana", "Cherry")

Dim idx As Variant
Sort a, idx                       ' by index: a is unchanged

Sort a
Dim pos As Variant
pos = Search(a, 4)                ' -> position of 4, or Null
```

### Objects with a custom `ICompare`

```vb
Dim people(1 To 3) As Object
Set people(1) = New Person: people(1).FirstName = "Charlie": people(1).Lastname = "Smith"
Set people(2) = New Person: people(2).FirstName = "Alice":   people(2).Lastname = "Jones"
Set people(3) = New Person: people(3).FirstName = "Bob":     people(3).Lastname = "Jones"

Set CompareCustom = New Person    ' Person implements ICompare
Sort people                       ' -> Jones Alice, Jones Bob, Smith Charlie
Set CompareCustom = Nothing       ' back to the automatic choice
```

---

## 🧠 How it works

| Phase | Algorithm | Condition |
|---|---|---|
| Large parts | Dual-Pivot QuickSort | 47 elements or more |
| Small parts | Insertion sort | fewer than 47 elements |

- **Two random pivots** from the left and right thirds; three parts: below the left pivot, between the pivots, above the right pivot.
- **Equal values**: when the middle part is larger than 4/7 of the array, the values equal to a pivot are moved to its ends, and a middle part between two equal pivots is not sorted again. Arrays with many equal values sort in O(n log n).
- **Recursive**: a readable showcase of interface-based polymorphism.
- **Descending** is native: the comparison order is passed down, there is no reverse pass.
- **Direction-aware search** reads the order from the first and last element.

---

## ⏱️ Performance

| Case | Before 2026-11-11 | Now |
|---|---|---|
| `Search`, 100,000 items | 2,488 µs | 4 µs |
| Sort 40,000 strings | 393 ms | 204 ms |
| Sort 40,000 random numbers | 227 ms | 168 ms |
| Sort 40,000 equal values | 15 ms | 11 ms |

---

## 📝 Changes

**2026-11-11** (the module was named `PolymorphicSort` in the former vba-sort-polymorphic repository)
- Fixed: sorting an array of objects without `CompareCustom` could crash the host. An object with a String default member (like `Person.FullName`) passed as a String, and was then swapped with `Let` instead of `Set`. Objects now raise error 17.
- Fixed: `Sort` by index of an array that is not 0-based raised error 9 (the index array was sorted between the bounds of the data array).
- Fixed: `Search` checked every element (and every index) on each call, O(n); it is a binary search again.
- New: dates in a Variant array are compared by operator (they raised error 17).
- Changed: the `ICompare` methods take their values `ByRef` (strings sort about twice as fast).
- Documented: an omitted `method` means `vbBinaryCompare`.
- Removed: three unused private functions.
- New: `Tests/TestPolySort.bas`.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
