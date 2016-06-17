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
// 3) Other nice things that a mail merge tool would need that isn't available in the tool
// described here and available elsewhere: https://iworkautomation.com/pages/script-tags-data-merge.html
// 4) Allow for find/replace text instead of the script tags (see #2)
// 5) Can we do Word/Excel as well? Do those even have AppleScript capability? (I haven't looked.)
// 6) Code organization/cleanup
// 7) Is there any real way to duplicate/copy a page without a keyboard shortcut in Pages?! (X_X)
// Unless we can AppleScript a way to manually select the first page to duplicate, we're kind of out
// of luck with a true mail merge in one document........
// Perhaps we can generate all the pages separately, then create a new document with all the pages?
// 8) Choose sheet (this only works with one sheet in Numbers since then the csv export makes a folder)
//   add an error if it can't find the file (error: more than 1 sheet!)
// 9) items with commas already in them come in with quotes in the strings >_>
// 10) refactor refactor refactor

// TODO: add alert on finished because duh

// Misc help:
// http://stackoverflow.com/questions/1968794/create-itunes-playlist-with-scripting-bridge
// https://developer.apple.com/library/mac/technotes/tn2084/_index.html

#import "ViewController.h"
#import "Pages.h"
#import "Numbers.h"
#import "CHCSVParser.h"
#import "PersonInfo.h"
#include <Carbon/Carbon.h>

@interface ViewController()

@property (weak) IBOutlet NSTextField *numbersInputPath;
@property (weak) IBOutlet NSTextField *pagesInputPath;


- (IBAction)chooseNumbersInputPath:(id)sender;
- (IBAction)choosePagesInputPath:(id)sender;

- (IBAction)startConvert:(id)sender;
- (IBAction)testNumbersColumns:(id)sender;
- (IBAction)testPagesScriptTags:(id)sender;
- (IBAction)testPagesDuplicate:(id)sender;
@property (unsafe_unretained) IBOutlet NSTextView *numbersColumns;
@property (unsafe_unretained) IBOutlet NSTextView *pagesTags;
@property (weak) IBOutlet NSProgressIndicator *activityIndicator;

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view.
    self.activityIndicator.hidden = YES;
}

- (void)setRepresentedObject:(id)representedObject {
    [super setRepresentedObject:representedObject];

    // Update the view, if already loaded.
}

-(void)startProgressIndicator {
    [self.activityIndicator startAnimation:self];
    self.activityIndicator.hidden = NO;
}

-(void)stopProgressIndicator {
    [self.activityIndicator stopAnimation:self];
    self.activityIndicator.hidden = YES;
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

-(void)showMessage:(NSString*)errorMessage {
    NSAlert *alert = [[NSAlert alloc] init];
    [alert addButtonWithTitle:@"OK"];
    [alert setMessageText:@"Info:"];
    [alert setInformativeText:errorMessage];
    [alert setAlertStyle:NSWarningAlertStyle];
    [alert runModal];
}

-(void)showError:(NSString*)errorMessage {
    NSAlert *alert = [[NSAlert alloc] init];
    [alert addButtonWithTitle:@"OK"];
    [alert setMessageText:@"Error!"];
    [alert setInformativeText:errorMessage];
    [alert setAlertStyle:NSWarningAlertStyle];
    [alert runModal];
}

// based on https://developer.apple.com/library/mac/technotes/tn2084/_index.html
-(void)runApplescriptForTag:(NSString*)tag withReplacementText:(NSString*)replacementText {
    NSString* path = [[NSBundle mainBundle] pathForResource:@"ReplaceItemInPages" ofType:@"applescript"];
    if (path != nil)
    {
        NSURL* url = [NSURL fileURLWithPath:path];
        if (url != nil)
        {
            NSDictionary* errors = [NSDictionary dictionary];
            NSAppleScript* appleScript = [[NSAppleScript alloc] initWithContentsOfURL:url error:&errors];
            if (appleScript != nil)
            {
                // create the first parameter
                NSAppleEventDescriptor* firstParameter = [NSAppleEventDescriptor descriptorWithString:tag];
                NSAppleEventDescriptor* secondParameter = [NSAppleEventDescriptor descriptorWithString:replacementText];
                
                // create and populate the list of parameters (in our case just one)
                NSAppleEventDescriptor* parameters = [NSAppleEventDescriptor listDescriptor];
                [parameters insertDescriptor:firstParameter atIndex:1]; // things are 1 indexed in AppleScript >_>
                [parameters insertDescriptor:secondParameter atIndex:2];
                
                // create the AppleEvent target
                ProcessSerialNumber psn = {0, kCurrentProcess};
                NSAppleEventDescriptor* target = [NSAppleEventDescriptor descriptorWithDescriptorType:typeProcessSerialNumber bytes:&psn
                                                                                               length:sizeof(ProcessSerialNumber)];
                
                // create an NSAppleEventDescriptor with the script's method name to call,
                // this is used for the script statement: "on show_message(user_message)"
                // Note that the routine name must be in lower case.
                NSAppleEventDescriptor* handler = [NSAppleEventDescriptor descriptorWithString: [@"replace_tag" lowercaseString]];
                
                // create the event for an AppleScript subroutine,
                // set the method name and the list of parameters
                NSAppleEventDescriptor* event = [NSAppleEventDescriptor appleEventWithEventClass:kASAppleScriptSuite eventID:kASSubroutineEvent
                                                                                targetDescriptor:target
                                                                                        returnID:kAutoGenerateReturnID
                                                                                   transactionID:kAnyTransactionID];
                [event setParamDescriptor:handler forKeyword:keyASSubroutineName];
                [event setParamDescriptor:parameters forKeyword:keyDirectObject];
                
                // call the event in AppleScript
                if (![appleScript executeAppleEvent:event error:&errors])
                {
                    // report any errors from 'errors'
                    NSLog(@"%@", errors);
                   // [self showError:@"ERroroero! %@"];
                }
            }
            else
            {
                // report any errors from 'errors'
                //[self showError:@"cowabunga!"];
            }
        }
    }
}

-(void)addPage {
    /*
    // This doesn't seem to work for adding a page. Hrm.
    PagesPage *page = [[[pages classForScriptingClass:@"page"] alloc] init];
    //[[pages.documents firstObject] insertObject:page atIndex:0];
    PagesDocument *docdoc = [[pages.documents firstObject] get];
    [[docdoc pages] insertObject:page atIndex:0];
     return;*/
    // TODO: error checking; return BOOL
    NSString* path = [[NSBundle mainBundle] pathForResource:@"DuplicatePage" ofType:@"applescript"];
    if (path != nil)
    {
        NSURL* url = [NSURL fileURLWithPath:path];
        if (url != nil)
        {
            NSDictionary* errors = [NSDictionary dictionary];
            NSAppleScript* appleScript = [[NSAppleScript alloc] initWithContentsOfURL:url error:&errors];
            if (appleScript != nil)
            {
                [appleScript executeAndReturnError:nil];
            }
        }
    }
}

- (IBAction)startConvert:(id)sender {
    if ([self.numbersInputPath.stringValue isEqualToString:@""])
        return;
    [self startProgressIndicator];
    dispatch_queue_t backgroundQueue = dispatch_queue_create("com.pikleproductions.mailmerge", NULL);
    
    dispatch_async(backgroundQueue, ^(void) {
        NumbersApplication *numbers = [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Numbers"];
        if (!numbers) {
            NSLog(@"No numbers :(");
            [self stopProgressIndicator];
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
            NSString *docsDir = [NSSearchPathForDirectoriesInDomains(NSApplicationSupportDirectory, NSUserDomainMask, YES) firstObject];
            NSString *appDir = [docsDir stringByAppendingPathComponent:@"NumbersToPagesMM"];
            
            if (![[NSFileManager defaultManager] fileExistsAtPath:appDir isDirectory:nil]) {
                [[NSFileManager defaultManager] createDirectoryAtPath:appDir withIntermediateDirectories:YES attributes:nil error:nil];
            }
            NSString *tmpPath = [appDir stringByAppendingPathComponent:@"tmp.csv"];
            [numbersDocument exportTo:[NSURL fileURLWithPath:tmpPath] as:NumbersExportFormatCSV withProperties:@{}];
            
            // Parse the CSV
            NSString *dataStr = [NSString stringWithContentsOfFile:tmpPath encoding:NSUTF8StringEncoding error:nil];
            if (!dataStr || [dataStr isEqualToString:@""]) {
                dispatch_async(dispatch_get_main_queue(), ^(void) {
                    [self showError:@"Couldn't get column names! Please only have 1 sheet in the Numbers document. This sheet may only have 1 table."];
                    [self stopProgressIndicator];
                });
                return;
            }
            NSArray *csvData = [dataStr CSVComponentsWithOptions:CHCSVParserOptionsUsesFirstLineAsKeys | CHCSVParserOptionsTrimsWhitespace | CHCSVParserOptionsSanitizesFields];
            if ([csvData count] == 0) {
                dispatch_async(dispatch_get_main_queue(), ^(void) {
                    [self showError:@"No data in Numbers file!"];
                    [self stopProgressIndicator];
                });
                return;
            }
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
            if (personInfo.count == 0) {
                dispatch_async(dispatch_get_main_queue(), ^(void) {
                    [self showError:@"Couldn't extract person info!"];
                    [self stopProgressIndicator];
                });
                return;
            }
            // Now insert into Pages!
            PagesApplication *pages = [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Pages"];
            if (!pages) {
                NSLog(@"No pages :(");
                [self stopProgressIndicator];
                return;
            }
            if ([pages isRunning])  {
                NSLog(@"Pages is already running");
                // TODO: see if file already open. if so, use PagesDocument *obj already open (?)
            }
            PagesDocument *pagesDocument = [[pages open:self.pagesInputPath.stringValue] get];
            if (pagesDocument) {
                SBElementArray<PagesPlaceholderText *> *placeholderTexts = [pagesDocument placeholderTexts];
                //            NSArray *items = [placeholderTexts get]; // returns strings
                unsigned long numToProcess = (unsigned long)[placeholderTexts count];
                // First, see if you need to duplicate pages. If you do, regrab the placeholder texts and such after duplicating.
                NSUInteger fieldsRequired = personInfo.count * [PersonInfo numFields];
                NSUInteger numberOfPersonInfoFields = [PersonInfo numFields];
                int tries = 0;
                while (fieldsRequired > numToProcess) {
                    [self addPage];
                    placeholderTexts = [pagesDocument placeholderTexts];
                    if (numToProcess == (unsigned long)[placeholderTexts count]) {
                        tries++;
                        if (tries == 5) {
                            dispatch_async(dispatch_get_main_queue(), ^(void) {
                                [self showError:@"Couldn't duplicate page! Did you remember to select it in Pages?"];
                                [self stopProgressIndicator];
                            });
                            return;
                        }
                    }
                    else {
                        tries = 0;
                    }
                    numToProcess = (unsigned long)[placeholderTexts count];
                    fieldsRequired = personInfo.count * numberOfPersonInfoFields;
                }
                // Can now start filling in the document!
                PersonInfo *currPersonInfo = [personInfo firstObject];
                NSUInteger currPersonInfoIndex = 0;
                NSUInteger numProcessed = 0;
                while (numProcessed < numToProcess) {
                    PagesPlaceholderText *text = [placeholderTexts firstObject];
                    if (currPersonInfo) {
                        // ...if I iterate through backwards, it works all the time. If I iterate through forwards,
                        // it stops returning valid items halfway through. (Could it be because the placeholders
                        // are no longer returned by the document after they've been edited, thus the array is getting "smaller"
                        // by 1 every time I replace a placeholder? Probably, especially given the return-references nature
                        // of the Scripting Bridge.
                        // If you want a for() loop [for (i in collection) loops do not work]:
                        // PagesPlaceholderText *text = [placeholderTexts objectAtIndex:numToProcess - j - 1];
                        if (text.tag && ![text.tag isEqualToString:@""]) {
                            //NSLog(@"Tag: %@", text.tag);
                            NSString *data = [currPersonInfo valueForTagKey:text.tag];
                            [self runApplescriptForTag:text.tag withReplacementText:data];
                        }
                        numProcessed++;
                        if (numProcessed % numberOfPersonInfoFields == 0) {
                            // Filled out all fields for a person (since we seem to be lucky and not get the next set of fields
                            // until we've filled out one set of fields
                            if (++currPersonInfoIndex <= personInfo.count - 1)
                                currPersonInfo = personInfo[currPersonInfoIndex];
                            else
                                currPersonInfo = nil;
                        }
                    }
                    else {
                        // Clear out all the other fields
                        [self runApplescriptForTag:text.tag withReplacementText:@""];
                        numProcessed++;
                    }
                }
            }
            else {
                dispatch_async(dispatch_get_main_queue(), ^(void) {
                    [self showError:@"Couldn't open the Pages document!"];
                    [self stopProgressIndicator];
                });
            }
        }
        else {
            dispatch_async(dispatch_get_main_queue(), ^(void) {
                [self showError:@"Couldn't open the Numbers document!"];
                [self stopProgressIndicator];
            });
        }
        dispatch_async(dispatch_get_main_queue(), ^(void) {
            [self stopProgressIndicator];
        });
    });
}

- (IBAction)testNumbersColumns:(id)sender {
    // TODO: refactor
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
        NSString *docsDir = [NSSearchPathForDirectoriesInDomains(NSApplicationSupportDirectory, NSUserDomainMask, YES) firstObject];
        NSString *appDir = [docsDir stringByAppendingPathComponent:@"NumbersToPagesMM"];
        
        if (![[NSFileManager defaultManager] fileExistsAtPath:appDir isDirectory:nil]) {
            [[NSFileManager defaultManager] createDirectoryAtPath:appDir withIntermediateDirectories:YES attributes:nil error:nil];
        }
        NSString *tmpPath = [appDir stringByAppendingPathComponent:@"tmp.csv"];
        if ([[NSFileManager defaultManager] fileExistsAtPath:tmpPath]) {
            [[NSFileManager defaultManager] removeItemAtPath:tmpPath error:nil];
        }
        [numbersDocument exportTo:[NSURL fileURLWithPath:tmpPath] as:NumbersExportFormatCSV withProperties:@{}];
        
        // Parse the CSV
        NSString *dataStr = [NSString stringWithContentsOfFile:tmpPath encoding:NSUTF8StringEncoding error:nil];
        if (!dataStr || [dataStr isEqualToString:@""]) {
            [self showError:@"Couldn't get column names! Please only have 1 sheet in the Numbers document. This sheet may only have 1 table."];
            return;
        }
        NSArray *csvData = [dataStr CSVComponentsWithOptions:CHCSVParserOptionsUsesFirstLineAsKeys];
        if ([csvData count] == 0) {
            [self showError:@"No data in Numbers file!"];
            return;
        }
        NSDictionary *dict = [csvData firstObject];
        NSMutableString *columnStr = [[NSMutableString alloc] initWithString:@""];
        int i = 1;
        for (NSString *key in dict)
            [columnStr appendString:[NSString stringWithFormat:@"%d: %@\n", i++, key]];
        self.numbersColumns.string = columnStr;
    }
}

- (IBAction)testPagesScriptTags:(id)sender {
    // TODO: refactor
    PagesApplication *pages = [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Pages"];
    if (!pages) {
        NSLog(@"No pages :(");
        return;
    }
    if ([pages isRunning])  {
        NSLog(@"Pages is already running");
        // TODO: see if file already open. if so, use PagesDocument *obj already open (?)
    }
    PagesDocument *pagesDocument = [[pages open:self.pagesInputPath.stringValue] get];
    if (pagesDocument) {
        SBElementArray<PagesPlaceholderText *> *placeholderTexts = [pagesDocument placeholderTexts];
        NSMutableArray *scriptTags = [[NSMutableArray alloc] init];
        for (PagesPlaceholderText *text in placeholderTexts) {
            if (![scriptTags containsObject:text.tag]) {
                [scriptTags addObject:text.tag];
            }
        }
        NSMutableString *pagesTagsStr = [[NSMutableString alloc] initWithString:@""];
        int i = 1;
        for (NSString *key in scriptTags)
            [pagesTagsStr appendString:[NSString stringWithFormat:@"%d: %@\n", i++, key]];
        self.pagesTags.string = pagesTagsStr;
    }
}

- (IBAction)testPagesDuplicate:(id)sender {
    [self addPage];
    [self showMessage:@"Check Pages to see if it duplicated a page. If not, it needs to be selected, and 'Duplicate Page' must be available under the edit menu for this to work."];
}

@end
