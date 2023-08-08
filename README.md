# What is poi-examples?
This shall be compilable examples for how to use [Apache POI](https://poi.apache.org/). Thats why using a GitHub repository.

It may contain usable tool classes too. But mainly it provides example code only.

# What is poi-examples not?
This is not a Java library to import and use via Jar library. Thats why it is not on Maven Central.

# How to run?
## Using git and maven

    git clone https://github.com/AxelRichter/poi-examples.git
    cd poi-examples
    mvn package -P copyDep

After that the directory `poi-examples` contains the `poi-examples-1.0-SNAPSHOT.jar` in `poi-examples/target` 
and the dependency libraries in `poi-examples/lib`. Thus:

    java -cp ./target/*;./target/lib/* arichter.examples.apache.poi.App
    
will run the GUI App.

## Whithout git and maven
Download the ZIP archive via Code - Download ZIP. Unzip that somewhere.

Then the `poi-examples-1.0-SNAPSHOT.jar` is in sub directory `run` 
and the dependency libraries in `run/lib`. Thus:

    java -cp ./run/*;./run/lib/* arichter.examples.apache.poi.App

will run the GUI App.
      

# Description of the single examples
See https://github.com/AxelRichter/poi-examples/wiki

# API Documentation
https://axelrichter.github.io/poi-examples/
