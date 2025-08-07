<div align="center">
  
  # hlang
  c style syntax language that compiles to VBA, written in VBA
</div>


## why
i never want to touch vba again.  

some environments only have excel.  

## Features / Improvements over VBA
- cleaner declaration syntax
- Consistent block syntax
- Shortcircuiting Logical operator (s)
- Upgraded VBA's foreach, now able to iterate over specific types instead of only```variant/object```
- Added ```break``` and ```continue```
- Removed need to add ```set``` when assigning objects
- Ability to create multiple classes in one file!
- Expression Statements
- Limited lambdas for use with ```map```, ```filter``` and ```reduce``` (not 100% tested)
- Slice operators for accessing array elements, slicing an array, and accessing excel values

## Development and Future Plans (not anytime soon)
Limitations: must be able to work entirely in a locked down environment with excel, word, powerpoint and chrome installed
- optional function arguments
- autobundle required functions
- include other files
- Create objects inplace: see ```functorcode.h```
- Closures
- Buffered Spreadsheet Read/Writes
- Type Checking
- better typing for arrays and dictionaries
- Try / Catch
- Better error handling than try/catch
- With statements
- array/dictionary initialisers
- inheritance
- fix impure iterables in foreach loops
- add variable resolver pass
- external dll support
- structs!
- tooling: basic editor with syntax highlighting, also able to run entirely in vba
- preprocessor macros
- bootstrap the language
- write a compiler in /compile to a non VBA language (while maintaining the original restrictions of compiling to VBA)


## TODO: 
- bootstrap 
-- done writing exprs
-- starting on parser , statement_f
- add inheritance (copy default implementations)
- import c libraries (for text editor implementation)
- variadic functions?
- class initialise too
- empty main files (used for classes only) will not make new files

## Use
```hcc <filename>``` in vba immediate window

cant find a way to get cli working without editing file associations
`src` - import target
`build` - export target

`output` - compile target

## Dev Env
create a blank `.xlsm` file. 
add references as shown below

```
Microsoft Scripting Runtime
Microsoft VBScript Regular Expressions 5.5
Microsoft Visual Basic for Applications Extensibility
```

drag in `import.bas`
run the sub `import`

now you have a master copy of the code, remove the code from `src` and `build`. (files should be committed to git anyway)
to backup at any point, export and commit
## Syntax
```
[one, two] => {} 
[one, two] => one + two


// dynamic arrays 

array hi;
arad(hi, item);

enum name {
  one, 
  two, 
  three

}
// no func needed within class
class Car < Vehicle{
  public void drive() {

  }

}


for (item : array) (String item) {
  int i;
  for (i : array2) (i) {
    pt(item, i);
  }
}
```
