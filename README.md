<a href="http://aimeos.org/">
    <img src="http://aimeos.org/fileadmin/template/icons/logo.png" alt="Aimeos logo" title="Aimeos" align="right" height="60" />
</a>

# Aimeos file container extension

[![Build Status](https://travis-ci.org/aimeos/ai-container.svg?branch=master)](https://travis-ci.org/aimeos/ai-container)
[![Coverage Status](https://coveralls.io/repos/aimeos/ai-container/badge.svg?branch=master)](https://coveralls.io/r/aimeos/ai-container?branch=master)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/aimeos/ai-container/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/aimeos/ai-container/?branch=master)
[![HHVM Status](http://hhvm.h4cc.de/badge/aimeos/ai-container.svg)](http://hhvm.h4cc.de/package/aimeos/ai-container)

The Aimeos container extension contains additonal container/content implementations for exporting and importing files. 
# Table of contents

- [Installation](#installation)
- [Usage](#usage)
  - [PHPExcel](#phpexcel)
- [License](#license)
- [Links](#links)

# Installation

As every Aimeos extension, the easiest way is to install it via [composer](https://getcomposer.org/). If you don't have composer installed yet, you can execute this string on the command line to download it:
```
php -r "readfile('https://getcomposer.org/installer');" | php -- --filename=composer
```

Add the container extension name to the "require" section of your ```composer.json``` (or your ```aimeos.composer.json```, depending on what is available) file:
```
"require": [
    ...
    "aimeos/ai-container": "dev-master"
],
```

Afterwards you only need to execute the composer update command on the command line:
```
composer update
```

If your composer file is named "aimeos.composer.json", you must use this:
```
COMPOSER=aimeos.composer.json composer update
```

These commands will install the Aimeos extension into the extension directory and it will be available immediately.

# Usage

Containers provide a single interface for handling container and content objects. They could be anything that can store one or more content objects (e.g. files) like directories, Zip files or PHPExcel documents. Content objects can be any binary or text file, CSV files or spread sheets.

There's a fine [documentation for working with containers](http://docs.aimeos.org/Developers/Utility/Create_and_read_files) available. The basic usage can be also found below.

**Export data to a container**
```
$container = MW_Container_Factory::getContainer( '/tmp/myfile', 'PHPExcel', 'PHPExcel', array() );

$content = $container->create( 'mysheet' );
$content->add( array( 'val1', 'val2', ... ) );
$container->add( $content );

$container->close();
```

**Read data from a container**
```
$container = MW_Container_Factory::getContainer( '/tmp/myfile.xls', 'PHPExcel', 'PHPExcel', array() );

foreach( $container as $content ) {
    foreach( $content as $data ) {
        print_r( $data );
    }
}

$container->close();
```

# License

The Aimeos container extension is licensed under the terms of the LGPLv3 Open Source license and is available for free.

# Links

* [Web site](http://aimeos.org/)
* [Documentation](http://docs.aimeos.org/)
* [Help](http://help.aimeos.org/)
* [Issue tracker](https://github.com/aimeos/ai-container/issues)
* [Source code](https://github.com/aimeos/ai-container)
