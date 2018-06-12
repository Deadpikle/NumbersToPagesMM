//
//  PersonInfo.h
//  NumbersToPagesMailMerge
//
//  Created by Deadpikle on 6/14/16.
//  Copyright © 2016 CIRC. All rights reserved.
//

#import <Foundation/Foundation.h>


// firstName, lastName, instrument, level, age, book, experience

// Use these for KVC
FOUNDATION_EXPORT NSString *const FirstNameColumnName;
FOUNDATION_EXPORT NSString *const LastNameColumnName;
FOUNDATION_EXPORT NSString *const InstrumentColumnName;
FOUNDATION_EXPORT NSString *const LevelColumnName;
FOUNDATION_EXPORT NSString *const AgeColumnName;
FOUNDATION_EXPORT NSString *const BookColumnName;
FOUNDATION_EXPORT NSString *const SightReadingColumnName;
FOUNDATION_EXPORT NSString *const ExperienceColumnName;

// These are used for matching up with the Pages document
FOUNDATION_EXPORT NSString *const FirstNameTagName;
FOUNDATION_EXPORT NSString *const LastNameTagName;
FOUNDATION_EXPORT NSString *const InstrumentTagName;
FOUNDATION_EXPORT NSString *const LevelTagName;
FOUNDATION_EXPORT NSString *const AgeTagName;
FOUNDATION_EXPORT NSString *const BookTagName;
FOUNDATION_EXPORT NSString *const SightReadingTagName;
FOUNDATION_EXPORT NSString *const ExperienceTagName;

// Everything is a string because that makes things easy. :3 (TODO: Maybe someday, don't use all strings.)
@interface PersonInfo : NSObject

@property NSString *firstName;
@property NSString *lastName;
@property NSString *instrument;
@property NSString *level;
@property NSString *age;
@property NSString *book;
@property NSString *sightReading;
@property NSString *experience;
@property NSString *city;
@property NSString *state;

@property NSMutableDictionary *unknownKeys;

-(id)initWithDictionary:(NSDictionary*)dict;

-(id)valueForTagKey:(NSString *)key;

@end
