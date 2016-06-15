//
//  PersonInfo.m
//  NumbersToPagesMailMerge
//
//  Created by Deadpikle on 6/14/16.
//  Copyright © 2016 CIRC. All rights reserved.
//

#import "PersonInfo.h"

NSString *const FirstNameColumnName = @"First Name";
NSString *const LastNameColumnName = @"Last Name";
NSString *const InstrumentColumnName = @"Instrument";
NSString *const LevelColumnName = @"Level";
NSString *const AgeColumnName = @"Age";
NSString *const BookColumnName = @"Book";
NSString *const SightReadingColumnName = @"Sight-read";
NSString *const ExperienceColumnName = @"Orch?";


NSString *const FirstNameTagName = @"First Name";
NSString *const LastNameTagName = @"Last Name";
NSString *const InstrumentTagName = @"Instrument";
NSString *const LevelTagName = @"Level";
NSString *const AgeTagName = @"Age";
NSString *const BookTagName = @"Book";
NSString *const SightReadingTagName = @"Sight-reading";
NSString *const ExperienceTagName = @"Orchestra"; // Experience ?

@implementation PersonInfo

-(id)initWithDictionary:(NSDictionary*)dict {
    self = [super init];
    if (self) {
        for (NSString *key in dict) {
            id value = [dict objectForKey:key];
            [self setValue:value forKey:key];
        }
    }
    return self;
}

+(NSUInteger)numFields {
    return 8;
}

-(void)setValue:(id)value forKey:(NSString *)key {
    if ([key isEqualToString:FirstNameColumnName]) {
        self.firstName = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:LastNameColumnName]) {
        self.lastName = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:InstrumentColumnName]) {
        self.instrument = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:LevelColumnName]) {
        self.level = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:AgeColumnName]) {
        self.age = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:BookColumnName]) {
        self.book = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:ExperienceColumnName]) {
        self.experience = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:SightReadingColumnName]) {
        self.sightReading = [NSString stringWithString:value];
    }
}

-(id)valueForKey:(NSString *)key {
    if ([key isEqualToString:FirstNameColumnName]) {
        return self.firstName;
    }
    else if ([key isEqualToString:LastNameColumnName]) {
        return self.lastName;
    }
    else if ([key isEqualToString:InstrumentColumnName]) {
        return self.instrument;
    }
    else if ([key isEqualToString:LevelColumnName]) {
        return self.level;
    }
    else if ([key isEqualToString:AgeColumnName]) {
        return self.age;
    }
    else if ([key isEqualToString:BookColumnName]) {
        return self.book;
    }
    else if ([key isEqualToString:ExperienceColumnName]) {
        return self.experience;
    }
    else if ([key isEqualToString:SightReadingColumnName]) {
        return self.sightReading;
    }
    return nil;
}

-(id)valueForTagKey:(NSString *)key {
    if ([key isEqualToString:FirstNameTagName]) {
        return self.firstName;
    }
    else if ([key isEqualToString:LastNameTagName]) {
        return self.lastName;
    }
    else if ([key isEqualToString:InstrumentTagName]) {
        return self.instrument;
    }
    else if ([key isEqualToString:LevelTagName]) {
        return self.level;
    }
    else if ([key isEqualToString:AgeTagName]) {
        return self.age;
    }
    else if ([key isEqualToString:BookTagName]) {
        return self.book;
    }
    else if ([key isEqualToString:ExperienceTagName]) {
        return self.experience;
    }
    else if ([key isEqualToString:SightReadingTagName]) {
        return self.sightReading;
    }
    return nil;
}
@end
