//
//  PersonInfo.h
//  NumbersToPagesMailMerge
//
//  Created by Deadpikle on 6/14/16.
//  Copyright © 2016 CIRC. All rights reserved.
//

#import <Foundation/Foundation.h>


// firstName, lastName, instrument, level, age, book, experience

@interface PersonInfo : NSObject

@property NSString *firstName;
@property NSString *lastName;
@property NSString *instrument;
@property NSString *level;
@property int age;
@property int book;
@property NSString *experience;

-(id)initWithDictionary:(NSDictionary*)dict;

@end
