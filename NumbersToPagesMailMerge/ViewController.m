//
//  ViewController.m
//  NumbersToPagesMailMerge
//
//  Created by Deadpikle on 6/14/16.
//  Copyright © 2016 CIRC. All rights reserved.
//

// TODO:
// 1) Allow user to choose which columns from CSV go into Pages
// 2) Allow user to set up tags for Pages (map columns to tags)
// 3) Other nice things

#import "ViewController.h"
#import "Pages.h"
#import "Numbers.h"
#import "CHCSVParser.h"
#import "PersonInfo.h"

@interface ViewController()

@property (weak) IBOutlet NSTextField *numbersInputPath;
@property (weak) IBOutlet NSTextField *pagesInputPath;

- (IBAction)chooseNumbersInputPath:(id)sender;
- (IBAction)choosePagesInputPath:(id)sender;

- (IBAction)startConvert:(id)sender;

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];

    // Do any additional setup after loading the view.
    NSLog(@"View did load!");
}

- (void)setRepresentedObject:(id)representedObject {
    [super setRepresentedObject:representedObject];

    // Update the view, if already loaded.
}

-(void)selectFileOfType:(NSString*)type forTextField:(NSTextField*)textField {
    // Create a File Open Dialog class.
    NSOpenPanel* panel = [NSOpenPanel openPanel];
    // Set array of file types
    NSArray *fileTypesArray = [NSArray arrayWithObjects:type, nil];
    // Enable options in the dialog.
    [panel setCanChooseFiles:YES];
    [panel setAllowedFileTypes:fileTypesArray];
    [panel setAllowsMultipleSelection:NO];
    [panel beginWithCompletionHandler:^(NSInteger result){
        if (result == NSFileHandlingPanelOKButton) {
            NSURL* path = [[panel URLs] objectAtIndex:0];
            // Open the document.
            textField.stringValue = path.path;
        }
    }];
}

- (IBAction)chooseNumbersInputPath:(id)sender {
    [self selectFileOfType:@"numbers" forTextField:self.numbersInputPath];
}

- (IBAction)choosePagesInputPath:(id)sender {
    [self selectFileOfType:@"pages" forTextField:self.pagesInputPath];
}

-(void)showError:(NSString*)errorMessage {
    NSAlert *alert = [[NSAlert alloc] init];
    [alert addButtonWithTitle:@"OK"];
    [alert setMessageText:@"Error!"];
    [alert setInformativeText:errorMessage];
    [alert setAlertStyle:NSWarningAlertStyle];
    [alert runModal];
}

- (IBAction)startConvert:(id)sender {
    if ([self.numbersInputPath.stringValue isEqualToString:@""])
        return;
    NumbersApplication *numbers = [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Numbers"];
    if (!numbers) {
        NSLog(@"No numbers :(");
        return;
    }
    if ([numbers isRunning])  {
        NSLog(@"Numbers is already running");
        // TODO: see if file already open. if so, use NumbersDocument *obj already open (?)
    }
    NumbersDocument *numbersDocument = [numbers open:self.numbersInputPath.stringValue];
    if (numbersDocument) {
        NSLog(@"Got the doc!");
        // NSAppleEventDescriptor *active = [NSAppleEventDescriptor descriptorWithEnumCode:OFProjectStatusActive];
        // Export the numbers data to CSV
        NSString *docsDir = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES) firstObject];
        NSString *appDir = [docsDir stringByAppendingPathComponent:@"NumbersToPagesMM"];
        
        if ([[NSFileManager defaultManager] fileExistsAtPath:appDir isDirectory:nil]) {
            [[NSFileManager defaultManager] createDirectoryAtPath:appDir withIntermediateDirectories:YES attributes:nil error:nil];
        }
        NSString *tmpPath = [appDir stringByAppendingPathComponent:@"tmp.csv"];
        [numbersDocument exportTo:[NSURL fileURLWithPath:tmpPath] as:NumbersExportFormatCSV withProperties:@{}];
        
        // Parse the CSV
        NSString *dataStr = [NSString stringWithContentsOfFile:tmpPath encoding:NSUTF8StringEncoding error:nil];
        NSArray *csvData = [dataStr CSVComponentsWithOptions:CHCSVParserOptionsUsesFirstLineAsKeys];
        //NSLog(@"Got csv data: %@", csvData);
        // firstName, lastName, instrument, level, age, book, experience
        NSMutableArray<PersonInfo*> *personInfo = [NSMutableArray array];
        for (NSDictionary *dict in csvData) {
            PersonInfo *info = [[PersonInfo alloc] initWithDictionary:dict];
            if (info) {
                // enforce first name, last name, age, and book
                if (![info.firstName isEqualToString:@""] &&
                    ![info.lastName isEqualToString:@""] &&
                    info.age != 0 &&
                    info.book != 0) {
                    [personInfo addObject:info];
                }
            }
        }
        // Now insert into Pages!
        PagesApplication *pages = [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Pages"];
        if (!pages) {
            NSLog(@"No pages :(");
            return;
        }
        if ([pages isRunning])  {
            NSLog(@"Pages is already running");
            // TODO: see if file already open. if so, use PagesDocument *obj already open (?)
        }
        
    }
    else {
        [self showError:@"Couldn't open the Numbers document!"];
    }
}

@end
