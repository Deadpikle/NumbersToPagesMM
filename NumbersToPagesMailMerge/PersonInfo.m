//
//  PersonInfo.m
//  NumbersToPagesMailMerge
//
//  Created by Deadpikle on 6/14/16.
//  Copyright © 2016 CIRC. All rights reserved.
//

#import "PersonInfo.h"

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

-(void)setValue:(id)value forKey:(NSString *)key {
    if ([key isEqualToString:@"firstName"]) {
        self.firstName = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:@"lastName"]) {
        self.lastName = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:@"instrument"]) {
        self.instrument = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:@"level"]) {
        self.level = [NSString stringWithString:value];
    }
    else if ([key isEqualToString:@"age"]) {
        self.age = [value intValue];
    }
    else if ([key isEqualToString:@"book"]) {
        self.book = [value intValue];
    }
    else if ([key isEqualToString:@"experience"]) {
        self.experience = [NSString stringWithString:value];
    }
}

@end
